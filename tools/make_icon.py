# -*- coding: utf-8 -*-
"""Generate assets/icon/docpilot.ico - the blue isometric cube app mark.

Drawn rather than traced so the shape stays crisp at every size Windows asks
for: the shell picks 16px for the taskbar and 256px for large tiles off the
same .ico, and a downscaled screenshot turns to mush at 16.

Everything is rendered at SUPER x the target and box-filtered down, which is
what gives the diagonals their clean anti-aliased edge.
"""
import os

from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "assets", "icon", "docpilot.ico")
SIZES = [256, 128, 64, 48, 32, 24, 16]
SUPER = 8

# top face is lit, left falls away, right sits between - the ordering is what
# reads as a solid object rather than three flat diamonds
TOP = ((0x6C, 0xB8, 0xF0), (0x4E, 0xA2, 0xE6))
LEFT = ((0x2B, 0x6F, 0xB4), (0x21, 0x59, 0x93))
RIGHT = ((0x3E, 0x8F, 0xD6), (0x2E, 0x74, 0xB8))


def _shade(draw, poly, top_rgb, bottom_rgb, box):
    """Fill a polygon with a vertical gradient, clipped to the polygon."""
    x0, y0, x1, y1 = box
    strip = Image.new("RGB", (1, max(1, y1 - y0)))
    px = strip.load()
    span = max(1, (y1 - y0) - 1)
    for y in range(y1 - y0):
        t = y / span
        px[0, y] = tuple(int(a + (b - a) * t) for a, b in zip(top_rgb, bottom_rgb))
    grad = strip.resize((max(1, x1 - x0), max(1, y1 - y0)))

    mask = Image.new("L", draw.im.size if hasattr(draw, "im") else (1, 1), 0)
    return grad, mask


def render(px):
    """One cube, rendered at px x px with a transparent ground."""
    S = px * SUPER
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

    # isometric half-width / half-height of the top diamond, and body depth
    cx = S / 2
    w = S * 0.385                     # leaves a margin so the shell can pad
    h = w * 0.5
    depth = w * 0.95                  # a true cube, not a slab
    cy = S / 2 - depth / 2 + h / 2

    top = [(cx, cy - h), (cx + w, cy), (cx, cy + h), (cx - w, cy)]
    left = [(cx - w, cy), (cx, cy + h), (cx, cy + h + depth), (cx - w, cy + depth)]
    right = [(cx + w, cy), (cx, cy + h), (cx, cy + h + depth), (cx + w, cy + depth)]

    for poly, (c0, c1) in ((left, LEFT), (right, RIGHT), (top, TOP)):
        xs = [p[0] for p in poly]
        ys = [p[1] for p in poly]
        box = (int(min(xs)), int(min(ys)), int(max(xs)) + 1, int(max(ys)) + 1)

        mask = Image.new("L", (S, S), 0)
        ImageDraw.Draw(mask).polygon(poly, fill=255)

        bw, bh = box[2] - box[0], box[3] - box[1]
        strip = Image.new("RGB", (1, bh))
        sp = strip.load()
        span = max(1, bh - 1)
        for y in range(bh):
            t = y / span
            sp[0, y] = tuple(int(a + (b - a) * t) for a, b in zip(c0, c1))
        grad = Image.new("RGB", (S, S), c1)
        grad.paste(strip.resize((bw, bh)), (box[0], box[1]))

        img.paste(grad, (0, 0), mask)

    return img.resize((px, px), Image.LANCZOS)


def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    frames = [render(s) for s in SIZES]
    # Pillow writes every requested size into the one .ico from the largest frame
    frames[0].save(OUT, format="ICO", sizes=[(s, s) for s in SIZES])
    frames[0].resize((512, 512), Image.LANCZOS).save(
        OUT.replace(".ico", ".png"), format="PNG")
    print(f"wrote {OUT} ({os.path.getsize(OUT)} bytes) sizes={SIZES}")


if __name__ == "__main__":
    main()
