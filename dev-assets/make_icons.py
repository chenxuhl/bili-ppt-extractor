# -*- coding: utf-8 -*-
"""生成 B站视频PPT提取器 的扩展图标（16/48/128px）。
粉底圆角矩形 + 白色幻灯片叠层图形，贴合 bilibili 品牌色。
"""
import os
from PIL import Image, ImageDraw

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "icons")
os.makedirs(OUT, exist_ok=True)

PINK = (251, 114, 153, 255)
PINK_DEEP = (255, 80, 117, 255)
WHITE = (255, 255, 255, 255)


def rounded_rect(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def make_icon(size: int) -> Image.Image:
    # 4x 超采样抗锯齿
    s = size * 4
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # 背景圆角矩形（垂直渐变近似：两段叠加）
    r = s // 5
    rounded_rect(d, [0, 0, s, s], r, PINK)
    rounded_rect(d, [0, s * 40 // 100, s, s], r, PINK_DEEP)
    # 重新画整体圆角，把下半渐变的方角裁掉：用遮罩合成更稳，这里直接叠一个上半透明带
    overlay = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    rounded_rect(od, [0, 0, s, s], r, PINK)
    img = Image.alpha_composite(img, overlay)
    d = ImageDraw.Draw(img)
    rounded_rect(d, [0, s * 42 // 100, s, s], r, PINK_DEEP)
    # 用圆角蒙版裁掉底边溢出的方角
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, s, s], radius=r, fill=255)
    img.putalpha(mask)

    # 幻灯片叠层：三张错开的白边框幻灯片
    w = s * 62 // 100   # 单张幻灯片宽
    h = int(w * 0.62)   # 16:10
    cx, cy = s // 2, s // 2
    off = s * 8 // 100
    lw = max(4, s // 32)

    # 后两张（错位）
    for dx, dy, alpha in ((-off * 2, off * 2, 110), (0, off, 170)):
        x0 = cx - w // 2 - off + dx
        y0 = cy - h // 2 - dy
        rr = s // 28
        fill = (255, 255, 255, 0)
        outline = (255, 255, 255, alpha)
        d.rounded_rectangle([x0, y0, x0 + w, y0 + h], radius=rr, fill=fill, outline=outline, width=lw)

    # 前一张（实心卡片）
    x0 = cx - w // 2 + off
    y0 = cy - h // 2
    rr = s // 28
    d.rounded_rectangle([x0, y0, x0 + w, y0 + h], radius=rr, fill=(255, 255, 255, 255))

    # 前景内容：标题条 + 柱状图
    pad = w // 7
    bx = x0 + pad
    bw = w - pad * 2
    # 标题条
    d.rounded_rectangle([bx, y0 + h // 6, bx + bw, y0 + h // 6 + h // 10], radius=h // 30, fill=PINK_DEEP)
    # 柱状图
    by = y0 + h // 2
    bar_h = h // 3
    bars = [(0, 45), (1, 100), (2, 70), (3, 55)]
    bar_w = bw // 9
    for i, pct in [b[:2] for b in bars]:
        bh = bar_h * pct // 100
        xbx = bx + i * (bar_w * 2)
        d.rounded_rectangle([xbx, by + bar_h - bh, xbx + bar_w, by + bar_h], radius=s // 90, fill=PINK)

    return img.resize((size, size), Image.LANCZOS)


for px in (16, 48, 128):
    make_icon(px).save(os.path.join(OUT, f"icon{px}.png"))
    print(f"icon{px}.png OK")
print("done ->", os.path.abspath(OUT))
