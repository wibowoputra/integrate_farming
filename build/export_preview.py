import argparse
import pathlib
import sys

import pythoncom
import win32com.client

BASE = pathlib.Path(__file__).resolve().parents[1]

ap = argparse.ArgumentParser()
ap.add_argument("--pptx", default=str(BASE / "output" / "integration_farmer_proposal.pptx"))
ap.add_argument("--preview", default=str(BASE / "build" / "preview"))
ap.add_argument("--report", default=str(BASE / "build" / "layout_report.txt"))
args = ap.parse_args()

PPTX = str(pathlib.Path(args.pptx).resolve())
PREVIEW = pathlib.Path(args.preview).resolve()
PREVIEW.mkdir(parents=True, exist_ok=True)
TOLERANCE_PT = 3.0

pythoncom.CoInitialize()
app = None
issues = []
try:
    app = win32com.client.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(PPTX, True, False, False)
    slide_w = pres.PageSetup.SlideWidth
    slide_h = pres.PageSetup.SlideHeight
    for i, slide in enumerate(pres.Slides, start=1):
        out = str(PREVIEW / f"slide_{i:02d}.png")
        slide.Export(out, "PNG", 1600, 900)
        for sh in slide.Shapes:
            name = sh.Name
            if sh.Left < -TOLERANCE_PT or sh.Top < -TOLERANCE_PT:
                issues.append(f"slide {i} '{name}': outside left/top "
                              f"(L={sh.Left:.0f}, T={sh.Top:.0f})")
            if sh.Left + sh.Width > slide_w + TOLERANCE_PT:
                issues.append(f"slide {i} '{name}': past right edge "
                              f"(L+W={sh.Left + sh.Width:.0f} > {slide_w:.0f})")
            if sh.Top + sh.Height > slide_h + TOLERANCE_PT:
                issues.append(f"slide {i} '{name}': past bottom edge "
                              f"(T+H={sh.Top + sh.Height:.0f} > {slide_h:.0f})")
            try:
                if sh.HasTextFrame and sh.TextFrame.HasText:
                    bh = sh.TextFrame2.TextRange.BoundHeight
                    if bh > sh.Height + TOLERANCE_PT:
                        sample = sh.TextFrame.TextRange.Text[:60].replace("\n", " ")
                        issues.append(
                            f"slide {i} '{name}': text overflow "
                            f"(bound={bh:.1f}pt > box={sh.Height:.1f}pt) :: {sample!r}"
                        )
            except Exception:
                pass
    pres.Close()
finally:
    if app is not None:
        try:
            app.Quit()
        except Exception:
            pass
    pythoncom.CoUninitialize()

report = pathlib.Path(args.report).resolve()
report.write_text(
    "\n".join(issues) + ("\n" if issues else "no layout issues\n"),
    encoding="utf-8",
)
print(f"Exported preview PNGs to {PREVIEW}")
if issues:
    print(f"LAYOUT ISSUES ({len(issues)}):")
    for msg in issues:
        print(" -", msg)
    sys.exit(1)
print("Layout validation PASSED")
