#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成应用图标 assets/icon.ico + icon.png（毛笔风"签"字）"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(exist_ok=True)
FONT = ROOT / "fonts" / "MaShanZheng-Regular.ttf"

S = 1024
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)

# 纸底圆角方块
d.rounded_rectangle([0, 0, S - 1, S - 1], radius=200, fill=(253, 251, 245, 255))
# 内描边
d.rounded_rectangle([14, 14, S - 15, S - 15], radius=188,
                    outline=(208, 198, 180, 255), width=6)

# 毛笔"签"字
glyph = Image.new("RGBA", (S, S), (0, 0, 0, 0))
gd = ImageDraw.Draw(glyph)
size = int(S * 0.74)
f = ImageFont.truetype(str(FONT), size)
bbox = gd.textbbox((0, 0), "签", font=f)
gw, gh = bbox[2] - bbox[0], bbox[3] - bbox[1]
gd.text(((S - gw) / 2 - bbox[0], (S - gh) / 2 - bbox[1]), "签", font=f,
        fill=(26, 26, 32, 255))
glyph = glyph.rotate(-6, resample=Image.BICUBIC, center=(S // 2, S // 2))
img.alpha_composite(glyph)

img.save(ASSETS / "icon.png")
img.save(ASSETS / "icon.ico",
         sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
print("icon written:", ASSETS / "icon.ico", (ASSETS / "icon.ico").stat().st_size, "bytes")
