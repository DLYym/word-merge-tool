# -*- coding: utf-8 -*-
"""生成应用图标 app.ico（蓝底 + 白色“合并”箭头图案）"""
from PIL import Image, ImageDraw

SIZE = 256


def draw_icon():
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # 圆角矩形背景
    r = 56
    d.rounded_rectangle([8, 8, SIZE - 8, SIZE - 8], radius=r, fill=(43, 125, 233, 255))

    # 两页文档（合并来源）
    d.rounded_rectangle([44, 74, 112, 182], radius=10, fill=(255, 255, 255, 235))
    d.rounded_rectangle([144, 74, 212, 182], radius=10, fill=(255, 255, 255, 235))
    # 文档内的文字横线
    for dx in (56, 156):
        d.rounded_rectangle([dx, 92, dx + 34, 100], radius=4, fill=(43, 125, 233, 255))
        d.rounded_rectangle([dx, 110, dx + 28, 118], radius=4, fill=(43, 125, 233, 160))
        d.rounded_rectangle([dx, 128, dx + 32, 136], radius=4, fill=(43, 125, 233, 160))

    # 中间的合并箭头
    d.rounded_rectangle([96, 122, 160, 132], radius=6, fill=(255, 255, 255, 255))
    d.polygon([(128, 150), (150, 128), (142, 128), (142, 112), (114, 112), (114, 128), (106, 128)],
              fill=(255, 255, 255, 255))

    return img


icon = draw_icon()
icon.save("/Users/a00/WorkBuddy/2026-09-02-08-47-18/word-merge-tool/app.ico",
          sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("icon 已生成")
