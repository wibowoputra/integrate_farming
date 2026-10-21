import pathlib
from svglib.svglib import svg2rlg
from reportlab.graphics import renderPM

BASE = pathlib.Path(__file__).resolve().parents[1]
SVG_DIR = BASE / "assets" / "svg"
OUT_DIR = BASE / "assets"

SIZES = {
    "cover_hero": (1400, 1000),
    "icon_sampah": (640, 640),
    "icon_dompet": (640, 640),
    "panen": (1200, 900),
    "pillar_sosial": (640, 640),
    "pillar_ekonomi": (640, 640),
    "pillar_lingkungan": (640, 640),
    "cycle_diagram": (2000, 800),
    "partnership": (900, 1400),
    "cta_hero": (1200, 900),
}

def convert(name, width, height):
    src = SVG_DIR / f"{name}.svg"
    dst = OUT_DIR / f"{name}.png"
    drawing = svg2rlg(str(src))
    if drawing is None:
        raise RuntimeError(f"failed to parse {src}")
    renderPM.drawToFile(
        drawing,
        str(dst),
        fmt="PNG",
        outputWidth=width,
        outputHeight=height,
    )
    print(f"OK {name} -> {dst.name} ({width}x{height})")

def main():
    failed = []
    for name, (w, h) in SIZES.items():
        try:
            convert(name, w, h)
        except Exception as e:
            failed.append((name, e))
            print(f"FAIL {name}: {e}")
    if failed:
        raise SystemExit(f"{len(failed)} conversion(s) failed")
    print("All SVG converted.")

if __name__ == "__main__":
    main()
