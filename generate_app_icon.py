from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


APP_ROOT = Path(__file__).resolve().parent
ICON_PATH = APP_ROOT / "H-Queex Finance App.ico"
PREVIEW_PATH = APP_ROOT / "H-Queex Finance App.png"
CANVAS_SIZE = 1024
NAVY = "#1f335d"
GOLD = "#b19154"
WHITE = "#f8f7f4"


def load_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        Path("C:/Windows/Fonts/georgiab.ttf"),
        Path("C:/Windows/Fonts/arialbd.ttf"),
        Path("C:/Windows/Fonts/segoeuib.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size=size)
    return ImageFont.load_default()


def main() -> None:
    image = Image.new("RGBA", (CANVAS_SIZE, CANVAS_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    inset = 72
    draw.rounded_rectangle(
        (inset, inset, CANVAS_SIZE - inset, CANVAS_SIZE - inset),
        radius=220,
        fill=NAVY,
    )

    font = load_font(470)
    draw.text((210, 180), "H", font=font, fill=WHITE)
    draw.text((470, 180), "Q", font=font, fill=WHITE)

    draw.rounded_rectangle((210, 470, 830, 555), radius=42, fill=GOLD)
    draw.rounded_rectangle((735, 435, 820, 760), radius=42, fill=GOLD)

    image.save(PREVIEW_PATH)
    image.save(ICON_PATH, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])


if __name__ == "__main__":
    main()