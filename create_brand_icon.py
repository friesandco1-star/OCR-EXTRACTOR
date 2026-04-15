#!/usr/bin/env python3

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def make_icon(path: Path) -> None:
    size = 256
    img = Image.new("RGBA", (size, size), (238, 238, 238, 255))
    draw = ImageDraw.Draw(img)

    source_logo = Path("assets") / "duck_logo.jpeg"
    if source_logo.exists():
        logo = Image.open(source_logo).convert("RGBA")
        logo.thumbnail((196, 196), Image.Resampling.LANCZOS)
        x = (size - logo.width) // 2
        y = (size - logo.height) // 2 - 8
        img.alpha_composite(logo, (x, y))
    else:
        draw.ellipse((52, 42, 204, 194), outline=(0, 0, 0, 255), width=8, fill=(255, 255, 255, 255))
        draw.ellipse((84, 104, 114, 134), fill=(0, 0, 0, 255))
        draw.ellipse((146, 92, 176, 122), fill=(0, 0, 0, 255))
        draw.rounded_rectangle((102, 102, 178, 164), radius=18, fill=(239, 143, 0, 255), outline=(0, 0, 0, 255), width=6)

    # Brand ribbon.
    draw.rounded_rectangle((36, 206, 220, 244), radius=12, fill=(0, 56, 188, 255))
    try:
        font = ImageFont.truetype("arialbd.ttf", 20)
    except Exception:
        font = ImageFont.load_default()
    draw.text((56, 216), "Sarang Tohokar", font=font, fill=(238, 238, 238, 255))

    path.parent.mkdir(parents=True, exist_ok=True)
    img.save(path, format="ICO", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])


def main() -> int:
    out = Path("assets") / "detailextract.ico"
    make_icon(out)
    print(f"Icon created: {out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
