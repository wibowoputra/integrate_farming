import pathlib
from PIL import Image, ImageDraw, ImageFont

BASE = pathlib.Path(__file__).resolve().parents[1]
OUT = BASE / "assets" / "v2" / "cycle_mosaic.jpg"

TILE_W, TILE_H = 660, 960
GAP = 6
BADGE_R = 46
MARGIN = 22

TILES = [
    ("icon_sampah.jpg", "1"),
    ("pillar_lingkungan.jpg", "2"),
    ("panen2.jpg", "3"),
    ("panen.jpg", "4"),
]


def cover_crop(im, w, h):
    im = im.convert("RGB")
    iw, ih = im.size
    ar = w / h
    if iw / ih > ar:
        nw = int(ih * ar)
        left = (iw - nw) // 2
        box = (left, 0, left + nw, ih)
    else:
        nh = int(iw / ar)
        top = (ih - nh) // 2
        box = (0, top, iw, top + nh)
    return im.crop(box).resize((w, h), Image.LANCZOS)


def font(size):
    for name in ("calibrib.ttf", "arialbd.ttf", "segoeuib.ttf"):
        p = pathlib.Path("C:/Windows/Fonts") / name
        if p.exists():
            return ImageFont.truetype(str(p), size)
    return ImageFont.load_default()


def main():
    total_w = TILE_W * len(TILES) + GAP * (len(TILES) - 1)
    canvas = Image.new("RGB", (total_w, TILE_H), (255, 255, 255))
    draw = ImageDraw.Draw(canvas)
    f = font(56)
    for i, (fname, num) in enumerate(TILES):
        src = BASE / "assets" / "v2" / fname
        tile = cover_crop(Image.open(src), TILE_W, TILE_H)
        x = i * (TILE_W + GAP)
        canvas.paste(tile, (x, 0))
        cx, cy = x + MARGIN + BADGE_R, MARGIN + BADGE_R
        draw.ellipse((cx - BADGE_R, cy - BADGE_R, cx + BADGE_R, cy + BADGE_R),
                     fill=(27, 94, 32))
        bbox = draw.textbbox((0, 0), num, font=f)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.text((cx - tw / 2, cy - th / 2 - 8), num, font=f, fill=(255, 255, 255))
    canvas.save(OUT, quality=90)
    print(f"Saved {OUT} ({canvas.size[0]}x{canvas.size[1]})")


if __name__ == "__main__":
    main()
