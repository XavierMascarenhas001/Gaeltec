"""
patch_page_template.py - one-off: adds the service-codes file and the
optional 'Variations_service codes' sheet to Tool 4 (Outputs) on the page.

Put this next to page_template.html and run it once:
    python patch_page_template.py
then rebuild as usual:
    python build_page.py
Running it a second time does nothing (it checks whether it's already applied).
"""
import os
import sys

OLD_DESC = ('const b3 = box(3, "Export Excel", "Summary (District › Project › Circuit pivot, click + to expand), '
            'Outputs (with total and orig), Variations (Source File › Job pivot with the most common comment, then the '
            'detail rows), one sheet per category and the CV8 / CV31 resumes. Drop the two logos once to have them in '
            'the top corner, as before.");')
NEW_DESC = ('const b3 = box(3, "Export Excel", "Summary (District › Project › Circuit pivot, click + to expand), '
            'Outputs (with total, orig and service code), Variations (Source File › Job pivot with Orig, Total, '
            'Variation and Reason, then the rows where variations isn\'t 0), one sheet per category and the CV8 / CV31 '
            'resumes. Drop the two logos once to have them in the top corner, as before.");')

OLD_BODY = '''  const nm = nameField("Output file name", "Outputs_Report.xlsx");
  const run = runArea("Export Excel", "save");
  b3.append(lr, h(`<div style="height:12px"></div>`), nm, h(`<div style="height:12px"></div>`), run);
  run.button.onclick = () => {
    if (!path) { run.fail("Load a file first (step 1)."); return; }
    run.run("outputs_export", {...payload(), cv_groups: cvBox.get(), output_name: nm.get(), logo_left: logoL.get(), logo_right: logoR.get()}, {needs: NEED_OUTPUTS, busy: "Generating the Excel file…"});
  };'''
NEW_BODY = '''  const scr = h(`<div class="row" style="margin-top:12px"></div>`);
  const scFile = dropField({label: "miscelaneous.parquet – service codes", kind: "file", types: [["Parquet", ".parquet"]],
                            hint: "optional – fills the “service code” column (Column_3, matched on the description)"});
  scr.append(scFile);
  const scSheet = h(`<label style="display:flex;gap:8px;align-items:center;margin-top:12px"><input type="checkbox"> Also add the <b>Variations_service codes</b> sheet (every service code with qty, qsub, Orig, Total and Variation)</label>`);
  const nm = nameField("Output file name", "Outputs_Report.xlsx");
  const run = runArea("Export Excel", "save");
  b3.append(lr, scr, scSheet, h(`<div style="height:12px"></div>`), nm, h(`<div style="height:12px"></div>`), run);
  run.button.onclick = () => {
    if (!path) { run.fail("Load a file first (step 1)."); return; }
    const wantSheet = $("input", scSheet).checked;
    if (wantSheet && !scFile.get()) { run.fail("Drop miscelaneous.parquet to add the Variations_service codes sheet."); return; }
    run.run("outputs_export", {...payload(), cv_groups: cvBox.get(), output_name: nm.get(), logo_left: logoL.get(), logo_right: logoR.get(),
                               misc: scFile.get(), service_code_sheet: wantSheet}, {needs: NEED_OUTPUTS, busy: "Generating the Excel file…"});
  };'''


def patch(text):
    if NEW_BODY in text:
        return text, False
    for old, new in ((OLD_DESC, NEW_DESC), (OLD_BODY, NEW_BODY)):
        if text.count(old) != 1:
            raise SystemExit("Couldn't find the Outputs section exactly as expected - nothing was changed.")
        text = text.replace(old, new)
    return text, True


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    target = sys.argv[1] if len(sys.argv) > 1 else os.path.join(here, "page_template.html")
    text = open(target, encoding="utf-8").read()
    text, changed = patch(text)
    if changed:
        open(target, "w", encoding="utf-8").write(text)
        print(f"Patched {os.path.basename(target)} - now run: python build_page.py")
    else:
        print(f"{os.path.basename(target)} already has the service-codes changes - nothing to do.")
