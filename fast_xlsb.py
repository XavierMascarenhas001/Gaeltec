"""fast_xlsb.py - a much faster drop-in for pandas' .xlsb reading.

pandas reads .xlsb through pyxlsb, which decodes every cell of every column
through several small Python objects, then pandas throws away the columns it
was not asked for. Control Files have ~90 columns of which Block1 needs 17,
so most of that work is wasted.

This reader walks the same BIFF12 records directly from memory and gives
pandas the *same* rows pandas would build from pyxlsb (same values, same
trimming/padding, same row numbers). Rows near the top (the header area) are
decoded in full; below that, only the requested columns get their values -
the others are only checked for "is this cell empty?", which is all pandas
uses them for before `usecols` drops them. Header handling, type inference
and usecols are still done by pandas itself, unchanged.

If anything unexpected turns up, `open_excel` falls back to plain pyxlsb.
"""
import struct
import zipfile
import xml.etree.ElementTree as ET

from pandas.io.excel._pyxlsb import PyxlsbReader

_I32 = struct.Struct("<i").unpack_from
_U32 = struct.Struct("<I").unpack_from
_D = struct.Struct("<d").unpack_from
_PACKD = struct.Struct("<d").unpack
_PACKU = struct.Struct("<I").pack

ROW, BLANK, NUM, BOOLERR, BOOL, FLOAT, STRING = 0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x07
F_STRING, F_FLOAT, F_BOOL, F_BOOLERR = 0x08, 0x09, 0x0A, 0x0B
CELL_IDS = {BLANK, NUM, BOOLERR, BOOL, FLOAT, STRING, F_STRING, F_FLOAT, F_BOOL, F_BOOLERR}   # 0x06 is ignored, as in pyxlsb
DIMENSION, SHEETDATA, SHEETDATA_END = 0x194, 0x191, 0x192
SHEET, SHEETS_END, SI, SST_END = 0x19C, 0x190, 0x13, 0x1A0


def _records(buf, pos=0):
    """(id, start, end) for each record - ids/lengths read exactly like pyxlsb."""
    n = len(buf)
    while pos < n:
        b = buf[pos]; pos += 1; rid = b
        if b & 0x80:
            for i in (1, 2, 3):
                if pos >= n:
                    return
                b = buf[pos]; pos += 1; rid += b << (8 * i)
                if not b & 0x80:
                    break
        if pos >= n:
            return
        b = buf[pos]; pos += 1; ln = b & 0x7F
        if b & 0x80:
            for i in (1, 2, 3):
                if pos >= n:
                    return
                b = buf[pos]; pos += 1; ln += (b & 0x7F) << (7 * i)
                if not b & 0x80:
                    break
        end = pos + ln
        if end > n:
            end = n
        yield rid, pos, end
        pos = end


def _wstr(buf, p, end):
    """pyxlsb read_string: int32 char count + UTF-16, None if the record is too short."""
    if p + 4 > end:
        return None
    ln = _U32(buf, p)[0]
    if p + 4 + ln * 2 > end:
        return None
    return bytes(buf[p + 4:p + 4 + ln * 2]).decode("utf-16", errors="replace")


def _rk(raw):
    """pyxlsb read_float (RK number)."""
    intval = raw if raw < 0x80000000 else raw - 0x100000000
    if intval & 0x02:
        v = float(intval >> 2)
    else:
        v = _PACKD(b"\x00\x00\x00\x00" + _PACKU(intval & 0xFFFFFFFC))[0]
    if intval & 0x01:
        v /= 100
    return v


class _Fallback(Exception):
    """Something unusual - let plain pyxlsb read this sheet instead."""


def _value(rid, buf, p, end, sst):
    """The value pyxlsb gives a cell record (p = just after column + style)."""
    if rid == NUM:
        return _rk(_U32(buf, p)[0]) if p + 4 <= end else None
    if rid == FLOAT or rid == F_FLOAT:
        return _D(buf, p)[0] if p + 8 <= end else None
    if rid == STRING:
        if p + 4 > end:
            raise _Fallback()
        i = _U32(buf, p)[0]
        return sst[i] if sst is not None else i
    if rid == F_STRING:
        return _wstr(buf, p, end)
    if rid == BOOL or rid == F_BOOL:
        return (buf[p] != 0) if p < end else True          # pyxlsb: None != 0
    if rid == BOOLERR or rid == F_BOOLERR:
        if p >= end:
            raise _Fallback()
        return hex(buf[p])
    return None                                                   # BLANK


def _convert(v):
    """pandas PyxlsbReader._convert_cell."""
    if v is None:
        return ""
    if isinstance(v, float):
        iv = int(v)
        return iv if iv == v else float(v)
    return v


class FastBook:
    def __init__(self, path):
        self._zf = zipfile.ZipFile(path)
        rels = {}
        with self._zf.open("xl/_rels/workbook.bin.rels") as fh:
            for el in ET.parse(fh).getroot():
                rels[el.attrib["Id"]] = el.attrib["Target"]
        buf = self._zf.read("xl/workbook.bin")
        self._sheets = []
        for rid, p, end in _records(buf):
            if rid == SHEET:
                q = p + 8
                relid = _wstr(buf, q, end)
                q += 4 + (len(relid) * 2 if relid is not None else 0)
                name = _wstr(buf, q, end)
                self._sheets.append((name, rels[relid]))
            elif rid == SHEETS_END:
                break
        self._sst = None
        self._sst_loaded = False

    @property
    def sheets(self):
        return [s for s, _ in self._sheets]

    def sst(self):
        if not self._sst_loaded:
            self._sst_loaded = True
            try:
                buf = self._zf.read("xl/sharedStrings.bin")
            except KeyError:
                return None
            out = []
            for rid, p, end in _records(buf):
                if rid == SI:
                    out.append(_wstr(buf, p + 1, end))
                elif rid == SST_END:
                    break
            self._sst = out
        return self._sst

    def get_sheet(self, idx):
        if isinstance(idx, str):
            idx = [s.lower() for s, _ in self._sheets].index(idx.lower()) + 1
        if idx < 1 or idx > len(self._sheets):
            raise IndexError("sheet index out of range")
        target = self._sheets[idx - 1][1].split("/")
        return FastSheet(self._zf.read("xl/{}/{}".format(target[0], target[-1])), self.sst())

    def close(self):
        self._zf.close()


class FastSheet:
    def __init__(self, buf, sst):
        self.buf = bytes(buf)
        self.sst = sst

    def data(self, need=None, full_rows=30, file_rows_needed=None):
        """Exactly pandas' PyxlsbReader.get_sheet_data(sheet) - except that, from
        row `full_rows` down, non-empty cells outside `need` come back as "".

        pandas pads every row to the widest row and keeps the gaps between rows,
        so for a column nobody asked for only two things can change the result:
        whether its row has any non-empty cell, and how wide the widest row is.
        Once both are already settled for a row, such cells are skipped unread.
        Cells must come in increasing column order within a row (Excel always
        writes them that way); anything else -> _Fallback (plain pyxlsb)."""
        buf, sst = self.buf, self.sst
        n = len(buf)
        u32 = _U32
        width = None
        data = []
        prev = -1
        cur_r = -1
        cells = None          # {col: value} - only non-empty cells of the current row
        lastc = -1            # last column seen in the current row
        gmax = -1             # widest non-empty column seen so far, all rows
        full = True
        started = False
        need = need if need is not None else None
        cell_ids = CELL_IDS
        pos = 0
        needl = None                          # needl[c]: is column c asked for?
        try:
            while pos < n:
                rid = buf[pos]
                # ---- fast path: a cell record (1-byte id 1..11, 1-byte length) ----
                if 0 < rid < 12 and rid != 6:
                    if buf[pos + 1] < 0x80:
                        p = pos + 2
                        end = p + buf[pos + 1]
                    else:                             # long record (e.g. a long formula text)
                        q = pos + 1
                        b = buf[q]; ln = b & 0x7F
                        for i in (1, 2, 3):
                            if not b & 0x80:
                                break
                            q += 1; b = buf[q]; ln += (b & 0x7F) << (7 * i)
                        p = q + 1
                        end = p + ln
                    pos = end
                    if end > n or cells is None or p + 8 > end:
                        raise _Fallback()
                    c = u32(buf, p)[0]
                    if c <= lastc or c >= width:
                        raise _Fallback()
                    lastc = c
                    if full or needl[c]:
                        v = _value(rid, buf, p + 8, end, sst)
                        if v is not None and not (v.__class__ is str and v == ""):
                            cells[c] = v
                            if c > gmax:
                                gmax = c
                        continue
                    if cells and c <= gmax:
                        continue                  # cannot change anything
                    q = p + 8
                    if rid == F_STRING:
                        empty = q + 4 > end or u32(buf, q)[0] == 0 or q + 4 + u32(buf, q)[0] * 2 > end
                    else:
                        v = _value(rid, buf, q, end, sst)
                        empty = v is None or (v.__class__ is str and v == "")
                    if not empty:
                        cells[c] = _SKIP
                        if c > gmax:
                            gmax = c
                    continue
                # ---- any other record: header read exactly like pyxlsb ----
                b = rid
                pos += 1
                if b & 0x80:
                    for i in (1, 2, 3):
                        b = buf[pos]; pos += 1; rid += b << (8 * i)
                        if not b & 0x80:
                            break
                b = buf[pos]; pos += 1; ln = b & 0x7F
                if b & 0x80:
                    for i in (1, 2, 3):
                        b = buf[pos]; pos += 1; ln += (b & 0x7F) << (7 * i)
                        if not b & 0x80:
                            break
                p = pos
                end = pos + ln
                if end > n:
                    raise _Fallback()
                pos = end
                if rid in cell_ids:
                    raise _Fallback()                 # cell record with a long header: unusual
                if rid == ROW:
                    if not started or p + 4 > end or width is None:
                        raise _Fallback()
                    r = u32(buf, p)[0]
                    if r != cur_r:
                        if cells:
                            row = [""] * (max(cells) + 1)
                            for c, v in cells.items():
                                if v is not _SKIP:
                                    row[c] = _convert(v)
                            data.extend([[]] * (cur_r - prev - 1))
                            data.append(row)
                            prev = cur_r
                            if file_rows_needed is not None and len(data) >= file_rows_needed:
                                return _pad(data)
                        cur_r = r
                        cells = {}
                        lastc = -1
                        full = need is None or r < full_rows
                    else:
                        raise _Fallback()             # repeated row record: let pyxlsb decide
                elif not started:
                    if rid == DIMENSION:
                        if p + 16 > end:
                            raise _Fallback()
                        width = u32(buf, p + 12)[0] + 1          # pyxlsb: dimension.c + dimension.w
                    elif rid == SHEETDATA:
                        started = True
                        if width is not None:
                            needl = [need is None or c in need for c in range(width)]
                elif rid == SHEETDATA_END:
                    if cells:
                        row = [""] * (max(cells) + 1)
                        for c, v in cells.items():
                            if v is not _SKIP:
                                row[c] = _convert(v)
                        data.extend([[]] * (cur_r - prev - 1))
                        data.append(row)
                    return _pad(data)
        except IndexError:
            raise _Fallback()
        raise _Fallback()                         # no SHEETDATA_END


_SKIP = object()


def _pad(data):
    if data:
        max_width = max(len(r) for r in data)
        if min(len(r) for r in data) < max_width:
            data = [r + (max_width - len(r)) * [""] for r in data]
    return data


class FastPyxlsbReader(PyxlsbReader):
    """pandas' own pyxlsb reader with the workbook/sheet access swapped for FastBook."""
    need = None

    def load_workbook(self, filepath_or_buffer, engine_kwargs):
        self.path = filepath_or_buffer
        return FastBook(filepath_or_buffer)

    def get_sheet_data(self, sheet, file_rows_needed=None):
        return sheet.data(self.need, 30, file_rows_needed)


def _letters_to_index(spec):
    """'A,B,AC,...' -> column indices (as pandas' usecols letters)."""
    out = []
    for part in str(spec).split(","):
        part = part.strip().upper()
        if ":" in part:
            a, b = part.split(":")
            out.extend(range(_col(a), _col(b) + 1))
        elif part:
            out.append(_col(part))
    return out


def _col(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n - 1


def parse(xl, sheet_name, usecols=None, **kw):
    """xl.parse(...) that tells a FastPyxlsbReader which columns it needs.
    Falls back to plain pyxlsb if the fast reader meets anything unusual."""
    rd = getattr(xl, "_reader", None)
    if isinstance(rd, FastPyxlsbReader):
        if usecols is None:
            rd.need = None
        elif isinstance(usecols, str):
            rd.need = set(_letters_to_index(usecols))
        elif all(isinstance(c, int) for c in usecols):
            rd.need = set(usecols)
        else:
            rd.need = None
        try:
            return xl.parse(sheet_name, usecols=usecols, **kw)
        except _Fallback:
            import pandas as pd
            return pd.read_excel(rd.path, sheet_name=sheet_name, usecols=usecols, engine="pyxlsb", **kw)
        finally:
            rd.need = None
    return xl.parse(sheet_name, usecols=usecols, **kw)


def open_excel(path):
    """pd.ExcelFile for any Control File; .xlsb gets the fast reader."""
    import pandas as pd
    if str(path).lower().endswith(".xlsb"):
        engines = pd.ExcelFile._engines
        orig = engines["pyxlsb"]
        try:
            engines["pyxlsb"] = FastPyxlsbReader
            xl = pd.ExcelFile(path, engine="pyxlsb")
            xl._reader.path = path            # for the plain-pyxlsb fallback in parse()
            return xl
        except Exception:
            engines["pyxlsb"] = orig
            return pd.ExcelFile(path, engine="pyxlsb")
        finally:
            engines["pyxlsb"] = orig
    return pd.ExcelFile(path)
