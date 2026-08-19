from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
PPTX = ROOT / "Files" / "PPT-Producer" / "Git使用过程短问题.pptx"
PHONE_IMAGE = Path("/tmp/network-reset-phone.png")
PHOTO_IMAGE = Path("/tmp/network-reset-instruction.png")
FONT_FILE = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"


def pil_font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_FILE, size)


def create_phone_screenshot() -> None:
    width, height = 768, 2048
    image = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(image)
    dark = "#25282D"
    muted = "#68707B"
    blue = "#1967D2"

    draw.text((26, 18), "12:32", font=pil_font(27), fill=dark)
    draw.text((615, 18), "HD  5G  ▪  21%  ⚡", font=pil_font(20), fill=dark)
    draw.text((30, 87), "☰", font=pil_font(35), fill=dark)
    draw.text((112, 88), "✣ 千问助手", font=pil_font(31), fill=blue)
    draw.text((351, 88), "千问⌄", font=pil_font(35), fill=dark)
    draw.text((620, 84), "⌁  ✧", font=pil_font(36), fill=dark)

    # Two small thumbnails at the top of the conversation.
    for x in (386, 570):
        draw.rounded_rectangle(
            (x, 174, x + 160, 338),
            radius=15,
            fill="#DDE2E5",
            outline="#B8C0C6",
            width=2,
        )
        draw.rectangle((x + 18, 198, x + 143, 313), fill="#EEF1F3")
        draw.line((x + 31, 230, x + 126, 230), fill="#D94A4A", width=4)
        draw.line((x + 31, 253, x + 125, 253), fill="#77818A", width=3)
        draw.line((x + 31, 276, x + 112, 276), fill="#77818A", width=3)

    draw.rounded_rectangle(
        (171, 376, 735, 458), radius=25, fill="#EAF3FF"
    )
    draw.text(
        (198, 398),
        "我没有找到网络重置的选项（win10）",
        font=pil_font(28),
        fill=dark,
    )

    body_lines = [
        "你当前在控制面板的“网络和 Internet”界面，但“网",
        "络重置”功能并不在传统控制面板中，而是位于",
        "Windows 设置 → 网络和 Internet → 高级网络设置",
        "路径下。",
        "",
        "具体操作步骤如下：",
        "",
        "• 点击左下角 开始 按钮，选择 设置（齿轮图标）。",
        "• 进入“网络和 Internet”。",
        "• 在左侧菜单中选择“高级网络设置”。",
        "• 向下滚动，找到“网络重置”选项并点击。",
        "• 点击“立即重置”，系统会提示你重新启动电脑。",
        "",
        "⚠ 注意：网络重置会清除所有网络适配器设置、",
        "Wi-Fi 密码、代理配置等，重启后需重新连接 Wi-Fi",
        "并输入密码。",
        "",
        "如果你希望快速访问该功能，也可以直接在“设置”顶",
        "部的搜索框输入“网络重置”，系统会自动定位到该",
        "选项。",
    ]
    y = 500
    for line in body_lines:
        draw.text((40, y), line, font=pil_font(27), fill=dark)
        y += 58 if line else 34

    # Bottom message controls.
    draw.line((0, 1785, width, 1785), fill="#ECEFF2", width=2)
    for x, label in [(40, "◉"), (132, "↗"), (224, "▢"), (316, "⌕")]:
        draw.rounded_rectangle(
            (x, 1815, x + 65, 1880),
            radius=16,
            fill="#FFFFFF",
            outline="#DFE3E7",
            width=2,
        )
        draw.text((x + 18, 1828), label, font=pil_font(27), fill=dark)
    draw.rounded_rectangle(
        (25, 1930, 740, 2016),
        radius=34,
        fill="#FFFFFF",
        outline="#E2E5E9",
        width=2,
    )
    draw.text(
        (103, 1954),
        "发消息或按住说话…",
        font=pil_font(25),
        fill="#B4BAC2",
    )
    image.save(PHONE_IMAGE, format="PNG", dpi=(200, 200), optimize=True)


def create_instruction_photo() -> None:
    width, height = 1332, 748
    image = Image.new("RGB", (width, height), "#CCD0CB")
    draw = ImageDraw.Draw(image)

    # Slight horizontal bands recreate the photographed-screen texture.
    for y in range(0, height, 4):
        color = "#C5CAC5" if (y // 4) % 2 == 0 else "#D1D5D0"
        draw.line((0, y, width, y), fill=color, width=1)

    dark = "#15253A"
    draw.text(
        (36, 99),
        "6. Windows「网络重置」（仍不行的…）",
        font=pil_font(52),
        fill=dark,
    )
    draw.rectangle((12, 224, 1129, 338), outline="#FF2638", width=10)
    draw.text(
        (67, 253),
        "设置 → 网络和 Internet → 高级网络设置 → 网络重置",
        font=pil_font(39),
        fill="#263A4B",
    )
    draw.text(
        (79, 358),
        "会清除网卡代理相关状态，重启后需重新连 Wi-Fi。",
        font=pil_font(32),
        fill="#425464",
    )
    draw.text(
        (79, 410),
        "重启后仍然先不要自动启动 FIClash。",
        font=pil_font(32),
        fill="#425464",
    )
    draw.text(
        (107, 532),
        "7. 阻止 FIClash「静默改回来」",
        font=pil_font(47),
        fill=dark,
    )
    draw.ellipse((1278, 212, 1390, 324), fill="#363636")
    draw.text((1305, 232), "›", font=pil_font(55), fill="#FFFFFF")
    image.save(PHOTO_IMAGE, format="PNG", dpi=(200, 200), optimize=True)


def add_text(
    slide,
    value: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    size: int,
    color: str,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
) -> None:
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    frame.margin_left = Inches(0.04)
    frame.margin_right = Inches(0.04)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = value
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)


def append_slide() -> None:
    prs = Presentation(PPTX)
    if len(prs.slides) != 6:
        raise RuntimeError(
            f"Expected the current six-slide Git questions deck, found {len(prs.slides)}"
        )

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string("F5F7FB")

    add_text(
        slide,
        "Windows 网络重置：操作记录",
        0.68,
        0.28,
        7.8,
        0.52,
        size=26,
        color="24324A",
        bold=True,
    )
    add_text(
        slide,
        "按照网络重置路径进行查找和操作",
        0.7,
        0.78,
        6.8,
        0.3,
        size=11,
        color="6B778C",
    )

    # Portrait screenshot, reduced and placed on the left.
    slide.shapes.add_picture(
        str(PHONE_IMAGE),
        Inches(0.74),
        Inches(1.2),
        width=Inches(3.0),
        height=Inches(5.56),
    )
    # Landscape photo, reduced and placed on the right.
    slide.shapes.add_picture(
        str(PHOTO_IMAGE),
        Inches(4.15),
        Inches(1.34),
        width=Inches(8.45),
        height=Inches(4.75),
    )

    badge = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(6.5),
        Inches(6.29),
        Inches(3.78),
        Inches(0.62),
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = RGBColor.from_string("FDECEC")
    badge.line.color.rgb = RGBColor.from_string("D94A4A")
    badge.line.width = Pt(1.5)
    add_text(
        slide,
        "没有操作成功",
        6.5,
        6.29,
        3.78,
        0.62,
        size=20,
        color="D94A4A",
        bold=True,
        align=PP_ALIGN.CENTER,
    )

    with NamedTemporaryFile(
        suffix=".pptx", dir=PPTX.parent, delete=False
    ) as temp_file:
        temp_path = Path(temp_file.name)
    try:
        prs.save(temp_path)
        temp_path.replace(PPTX)
    finally:
        temp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    create_phone_screenshot()
    create_instruction_photo()
    append_slide()
    print(PPTX)
