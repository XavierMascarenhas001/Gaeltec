"""fast_xlsx.py - write DataFrames to .xlsx many times faster than pandas.to_excel.

The workbook holds the same cells as `df.to_excel(path, index=False)` with the
xlsxwriter engine (the one pandas picks by default): header row bold, bordered
and centred; dates as Excel serial numbers shown "YYYY-MM-DD HH:MM:SS"; empty
values left empty; "=..." text written as a formula; links, long text and the
1900 date quirks handled the same way. Reading it back with pandas gives the
same DataFrame as the pandas/xlsxwriter file.

It is faster because it builds the sheet XML column by column, instead of
pandas creating a cell object (plus a style lookup) for every single cell.
"""
import datetime as _dt
import re
import zipfile
from decimal import Decimal

import numpy as np
import pandas as pd

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="YYYY-MM-DD HH:MM:SS"/><numFmt numFmtId="165" formatCode="YYYY-MM-DD"/></numFmts><fonts count="3"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><b/><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font><font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/></styleSheet>"""
S_HEAD, S_DATETIME, S_DATE, S_LINK, S_DAYS = 1, 2, 3, 4, 5

_EPOCH_NS = int(np.datetime64("1899-12-31T00:00:00", "ns").astype("int64"))
_DAY_NS = 86400 * 10**9
_STRMAX = 32767
_URLMAX = 2079
_URL_RE = re.compile(r"(?:(?:ftp|http)s?://|mailto:|(?:in|ex)ternal:|file://)")
_XML_ESC = re.compile('["&<>\n]')
_CTRL_1 = re.compile("(_x[0-9a-fA-F]{4}_)")
_CTRL_2 = re.compile(r"([\x00-\x08\x0b-\x1f])")
_LEAD_WS = re.compile(r"^\s")
_TRAIL_WS = re.compile(r"\s$")
_INF, _NINF = float("inf"), float("-inf")


def _col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _esc(s):
    """xlsxwriter's _escape_data."""
    if not _XML_ESC.search(s):
        return s
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _si(s):
    """One <si> element exactly as xlsxwriter's shared-strings writer makes it."""
    s = _CTRL_1.sub(r"_x005F\1", s)
    s = _CTRL_2.sub(lambda m: f"_x{ord(m.group(1)):04X}_", s)
    s = s.replace("￾", "_xFFFE_").replace("￿", "_xFFFF_")
    if s.startswith("<r>") and s.endswith("</r>"):
        return f"<si>{s}</si>"
    sp = ' xml:space="preserve"' if (_LEAD_WS.search(s) or _TRAIL_WS.search(s)) else ""
    return f"<si><t{sp}>{_esc(s)}</t></si>"


def _serial_ns(ns):
    """xlsxwriter's datetime -> Excel serial for an int64 array of datetime64[ns]:
    days since 1899-12-31 + time of day to the microsecond, -1 on 1900-01-01
    and +1 after 1900-02-28 (Excel's leap-year bug) - same float order."""
    rel = ns - _EPOCH_NS
    days = rel // _DAY_NS
    in_day = rel - days * _DAY_NS
    secs = in_day // 10**9
    us = (in_day - secs * 10**9) // 1000
    t = days.astype("float64") + (secs.astype("float64") + us.astype("float64") / 1e6) / 86400
    t = np.where(days == 1, t - 1, t)
    return np.where(t > 59, t + 1, t)


def _serial_one(v):
    """Same as _serial_ns for one datetime / Timestamp / date (xlsxwriter code)."""
    is_dt = isinstance(v, _dt.datetime)
    if not is_dt:
        v = _dt.datetime.fromordinal(v.toordinal())
    delta = v - _dt.datetime(1899, 12, 31)
    t = delta.days + (float(delta.seconds) + float(delta.microseconds) / 1e6) / (60 * 60 * 24)
    if is_dt and v.isocalendar() == (1900, 1, 1):
        t -= 1
    if t > 59:
        t += 1
    return t


class _Book:
    def __init__(self):
        self.idx, self.items = {}, []
        self.links = 0

    def sst(self, s):
        i = self.idx.get(s)
        if i is None:
            i = self.idx[s] = len(self.items)
            self.items.append(s)
        return i


def _text_cell(ref, v, book):
    """A str value the way xlsxwriter's write() handles it."""
    if v == "":
        return ""
    if v[0] == "=" or (v.startswith("{=") and v.endswith("}")):
        f = v[1:] if v[0] == "=" else v[2:-1]
        return f'<c r="{ref}"><f>{_esc(f)}</f><v>0</v></c>'
    if ":" in v and _URL_RE.match(v):
        url, _, anchor = v.partition("#")
        if len(url) > _URLMAX or len(anchor) > _URLMAX:
            return ""                                    # xlsxwriter ignores the link
        book.links += 1
        if book.links > 65530:
            return ""                                    # over Excel's link limit: ignored
        try:
            from xlsxwriter.url import Url
            text = Url(v).text
        except Exception:
            text = v
        if len(text) > _STRMAX:
            text = text[:_STRMAX]
        return f'<c r="{ref}" s="{S_LINK}" t="s"><v>{book.sst(text)}</v></c>'
    if len(v) > _STRMAX:
        v = v[:_STRMAX]
    return f'<c r="{ref}" t="s"><v>{book.sst(v)}</v></c>'


def _obj_cell(ref, v, book):
    """Any other value, the way pandas' ExcelFormatter + xlsxwriter write it."""
    if v is None:
        return ""
    try:
        if pd.api.types.is_scalar(v) and pd.isna(v):
            return ""                                                 # na_rep ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (bool, np.bool_)):
        return f'<c r="{ref}" t="b"><v>{int(bool(v))}</v></c>'
    if isinstance(v, (int, np.integer)):
        return f'<c r="{ref}"><v>{int(v):.16G}</v></c>'
    if isinstance(v, (float, np.floating)):
        v = float(v)
        if v == _INF or v == _NINF:
            return _text_cell(ref, "inf" if v > 0 else "-inf", book)  # pandas inf_rep
        return f'<c r="{ref}"><v>{v:.16G}</v></c>'
    if isinstance(v, np.datetime64):
        v = pd.Timestamp(v)
    if isinstance(v, _dt.datetime):                                   # incl. pd.Timestamp
        if v.tzinfo is not None:
            raise ValueError("Excel does not support datetimes with timezones.")
        return f'<c r="{ref}" s="{S_DATETIME}"><v>{_serial_one(v):.16G}</v></c>'
    if isinstance(v, _dt.date):
        return f'<c r="{ref}" s="{S_DATE}"><v>{_serial_one(v):.16G}</v></c>'
    if isinstance(v, (_dt.timedelta, np.timedelta64)):
        days = pd.Timedelta(v).total_seconds() / 86400
        return f'<c r="{ref}" s="{S_DAYS}"><v>{days:.16G}</v></c>'
    if isinstance(v, Decimal):
        return f'<c r="{ref}"><v>{v:.16G}</v></c>'
    return _text_cell(ref, str(v), book)


def _column_cells(series, letter, book, first_row):
    """['<c .../>' or '' for each value] for one column; the first value goes to
    Excel row `first_row`."""
    vals = series.to_numpy()
    rows = range(first_row, first_row + len(vals))
    kind = series.dtype.kind
    if kind == "M":                                                   # datetime64
        if getattr(series.dt, "tz", None) is not None:
            raise ValueError("Excel does not support datetimes with timezones.")
        d = vals.astype("datetime64[ns]")
        nat = np.isnat(d).tolist()
        ser = _serial_ns(d.astype("int64")).tolist()
        return ["" if na else f'<c r="{letter}{r}" s="{S_DATETIME}"><v>{v:.16G}</v></c>' for r, v, na in zip(rows, ser, nat)]
    if kind == "f":
        fv = vals.astype("float64")
        bad = (~np.isfinite(fv)).tolist()
        out = []
        for r, v, b in zip(rows, fv.tolist(), bad):
            if not b:
                out.append(f'<c r="{letter}{r}"><v>{v:.16G}</v></c>')
            elif v != v:
                out.append("")
            else:
                out.append(_text_cell(f"{letter}{r}", "inf" if v > 0 else "-inf", book))
        return out
    if kind in "iu":
        return [f'<c r="{letter}{r}"><v>{v:.16G}</v></c>' for r, v in zip(rows, vals.tolist())]
    if kind == "b":
        return [f'<c r="{letter}{r}" t="b"><v>{int(v)}</v></c>' for r, v in zip(rows, vals.tolist())]
    if kind != "O":
        return [_obj_cell(f"{letter}{r}", v, book) for r, v in zip(rows, series.tolist())]
    out = []
    ap = out.append
    for r, v in zip(rows, vals):
        if v.__class__ is str:                                        # the common case first
            ap(_text_cell(f"{letter}{r}", v, book))
        else:
            ap(_obj_cell(f"{letter}{r}", v, book))
    return out


def _sheet_xml_chunks(df, book, chunk=20000):
    n, ncol = len(df), len(df.columns)
    letters = [_col_letter(i) for i in range(ncol)]
    last = letters[-1] if ncol else "A"
    yield ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           f'<dimension ref="A1:{last}{n + 1}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>')
    head = "".join(f'<c r="{letters[i]}1" s="{S_HEAD}" t="s"><v>{book.sst(str(c))}</v></c>' for i, c in enumerate(df.columns))
    yield f'<row r="1">{head}</row>'
    for start in range(0, n, chunk):
        part = df.iloc[start:start + chunk]
        r0 = start + 2
        cols = [_column_cells(part.iloc[:, i], letters[i], book, r0) for i in range(ncol)]
        if cols:
            yield "".join(f'<row r="{r0 + j}">{"".join(rc)}</row>' for j, rc in enumerate(zip(*cols)))
    yield '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>'


def write(path, sheets):
    """sheets: [(sheet_name, DataFrame)] -> one .xlsx file (index not written)."""
    book = _Book()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        names = []
        for k, (name, df) in enumerate(sheets, start=1):
            names.append(name)
            book.links = 0                                # the link limit is per sheet
            with z.open(f"xl/worksheets/sheet{k}.xml", "w", force_zip64=True) as fh:
                for piece in _sheet_xml_chunks(df, book):
                    fh.write(piece.encode("utf-8"))
        with z.open("xl/sharedStrings.xml", "w", force_zip64=True) as fh:
            fh.write(('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                      f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(book.items)}" uniqueCount="{len(book.items)}">').encode("utf-8"))
            for i in range(0, len(book.items), 5000):
                fh.write("".join(_si(s) for s in book.items[i:i + 5000]).encode("utf-8"))
            fh.write(b"</sst>")
        z.writestr("xl/styles.xml", STYLES)
        sheets_xml = "".join(f'<sheet name="{_esc(n)}" sheetId="{k}" r:id="rId{k}"/>' for k, n in enumerate(names, start=1))
        z.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView activeTab="0"/></bookViews>'
                   f"<sheets>{sheets_xml}</sheets></workbook>")
        rels = "".join(f'<Relationship Id="rId{k}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{k}.xml"/>' for k in range(1, len(names) + 1))
        k = len(names)
        rels += (f'<Relationship Id="rId{k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                 f'<Relationship Id="rId{k + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + "</Relationships>")
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>')
        now = _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        z.writestr("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                   'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                   f'<dc:creator>Gaeltec Tools</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>')
        z.writestr("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>')
        ov = "".join(f'<Override PartName="/xl/worksheets/sheet{k}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for k in range(1, len(names) + 1))
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + ov +
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                   '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                   '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')
    return path
