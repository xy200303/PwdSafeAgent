from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "build"
OUT_DIR.mkdir(exist_ok=True)

SIZE = 1024
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

bg = (13, 45, 78, 255)
accent = (34, 116, 165, 255)
gold = (221, 169, 64, 255)
white = (247, 250, 252, 255)

draw.rounded_rectangle((72, 72, SIZE - 72, SIZE - 72), radius=220, fill=bg)
draw.rounded_rectangle((114, 114, SIZE - 114, SIZE - 114), radius=178, outline=(88, 139, 174, 255), width=24)
draw.polygon(
    [
        (SIZE / 2, 226),
        (724, 318),
        (684, 704),
        (SIZE / 2, 822),
        (340, 704),
        (300, 318),
    ],
    fill=accent,
)
draw.line([(352, 330), (SIZE / 2, 254), (672, 330)], fill=(122, 177, 211, 255), width=18)

try:
    font = ImageFont.truetype("arialbd.ttf", 380)
except OSError:
    font = ImageFont.load_default()

text = "P"
box = draw.textbbox((0, 0), text, font=font)
text_w = box[2] - box[0]
text_h = box[3] - box[1]
draw.text(((SIZE - text_w) / 2, 330 - text_h / 2), text, font=font, fill=white)

draw.rounded_rectangle((408, 616, 616, 748), radius=48, fill=gold)
draw.arc((432, 520, 592, 690), 180, 360, fill=gold, width=42)
draw.ellipse((493, 666, 531, 704), fill=bg)

png_path = OUT_DIR / "icon.png"
ico_path = OUT_DIR / "icon.ico"
img.save(png_path)
img.save(ico_path, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])

print(f"Generated {png_path}")
print(f"Generated {ico_path}")
