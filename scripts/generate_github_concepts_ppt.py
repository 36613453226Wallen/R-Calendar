from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "Files" / "PPT-Producer"
IMAGE_DIR = OUTPUT_DIR / "Teaching-Associate教授の助理 职位"
PPTX_PATH = OUTPUT_DIR / "GitHub三大核心概念.pptx"
REFERENCE_PATH = IMAGE_DIR / "Teaching-Associate职位对照表.png"

FONT_CN = "Microsoft YaHei"
FONT_EN = "Aptos"
FONT_FILE = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"

NAVY = "152238"
NAVY_2 = "20324F"
WHITE = "FFFFFF"
BG = "F5F7FB"
TEXT = "26344A"
MUTED = "66758A"
LINE = "DCE3ED"
BLUE = "2F6BFF"
BLUE_LIGHT = "EAF0FF"
PURPLE = "7757E8"
PURPLE_LIGHT = "F0EDFF"
GREEN = "13A36D"
GREEN_LIGHT = "E5F6EF"
ORANGE = "F59B42"
ORANGE_LIGHT = "FFF1E3"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def add_text(
    slide,
    text: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    size: int = 20,
    color: str = TEXT,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.MIDDLE,
    font_name: str = FONT_CN,
    margin: float = 0.04,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = valign
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_after = Pt(0)
    run = paragraph.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return box


def add_round_rect(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill: str = WHITE,
    line: str = LINE,
    radius_shape=MSO_SHAPE.ROUNDED_RECTANGLE,
):
    shape = slide.shapes.add_shape(
        radius_shape, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line)
    shape.line.width = Pt(1)
    return shape


def add_connector(
    slide,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    color: str = LINE,
    width: float = 2,
):
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1),
        Inches(y1),
        Inches(x2),
        Inches(y2),
    )
    connector.line.color.rgb = rgb(color)
    connector.line.width = Pt(width)
    return connector


def set_background(slide, color: str) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def add_slide_title(slide, index: str, title: str, subtitle: str) -> None:
    add_text(
        slide,
        index,
        0.65,
        0.42,
        0.45,
        0.38,
        size=13,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(0.62), Inches(0.38), Inches(0.52), Inches(0.52)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = rgb(BLUE)
    badge.line.fill.background()
    # Keep the text badge above the oval.
    slide.shapes._spTree.remove(slide.shapes._spTree[-2])
    slide.shapes._spTree.insert(len(slide.shapes._spTree), badge._element)
    add_text(
        slide,
        index,
        0.62,
        0.38,
        0.52,
        0.52,
        size=13,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_text(slide, title, 1.28, 0.26, 7.9, 0.48, size=26, bold=True)
    add_text(slide, subtitle, 1.3, 0.76, 10.8, 0.3, size=11, color=MUTED)
    add_connector(slide, 0.65, 1.18, 12.68, 1.18, color=LINE, width=1)


def add_info_card(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    label: str,
    body: str,
    color: str,
    light: str,
) -> None:
    add_round_rect(slide, x, y, w, h, fill=WHITE, line=LINE)
    add_round_rect(slide, x + 0.22, y + 0.2, 0.86, 0.34, fill=light, line=light)
    add_text(
        slide,
        label,
        x + 0.22,
        y + 0.2,
        0.86,
        0.34,
        size=11,
        color=color,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_text(
        slide,
        body,
        x + 0.22,
        y + 0.66,
        w - 0.44,
        h - 0.82,
        size=15,
        color=TEXT,
        valign=MSO_ANCHOR.TOP,
    )


def add_footer(slide, page: int) -> None:
    add_text(
        slide,
        "GitHub 协作与自动化 · 概念速览",
        0.65,
        7.15,
        5.8,
        0.18,
        size=9,
        color=MUTED,
    )
    add_text(
        slide,
        f"{page:02d}",
        12.25,
        7.12,
        0.42,
        0.2,
        size=9,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )


def title_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, NAVY)
    add_round_rect(slide, 0.72, 0.62, 1.66, 0.36, fill=NAVY_2, line=NAVY_2)
    add_text(
        slide,
        "GITHUB · CONCEPTS",
        0.72,
        0.62,
        1.66,
        0.36,
        size=10,
        color="AFC2E4",
        bold=True,
        align=PP_ALIGN.CENTER,
        font_name=FONT_EN,
    )
    add_text(
        slide,
        "GitHub 三大核心概念",
        0.72,
        1.35,
        8.5,
        0.78,
        size=36,
        color=WHITE,
        bold=True,
    )
    add_text(
        slide,
        "从代码协作，到后台自动化，再到云端交付",
        0.75,
        2.16,
        8.5,
        0.48,
        size=18,
        color="B7C5DC",
    )

    cards = [
        (0.75, BLUE, "01", "Pull Request", "代码审查与合并"),
        (4.56, PURPLE, "02", "Actions Runner", "后台执行自动化任务"),
        (8.37, GREEN, "03", "云端集成", "连接部署、通知与服务"),
    ]
    for x, color, number, heading, subheading in cards:
        add_round_rect(slide, x, 3.42, 3.36, 2.05, fill=NAVY_2, line="3A4C68")
        add_round_rect(slide, x + 0.25, 3.68, 0.54, 0.54, fill=color, line=color)
        add_text(
            slide,
            number,
            x + 0.25,
            3.68,
            0.54,
            0.54,
            size=13,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            heading,
            x + 0.25,
            4.38,
            2.8,
            0.42,
            size=21,
            color=WHITE,
            bold=True,
            font_name=FONT_EN if number != "03" else FONT_CN,
        )
        add_text(
            slide,
            subheading,
            x + 0.25,
            4.84,
            2.8,
            0.3,
            size=13,
            color="B7C5DC",
        )
    add_connector(slide, 4.12, 4.45, 4.5, 4.45, color="637894", width=2)
    add_connector(slide, 7.93, 4.45, 8.31, 4.45, color="637894", width=2)
    add_text(
        slide,
        "概念说明 · 场景示例 · 协同流程",
        0.75,
        6.78,
        5.0,
        0.25,
        size=11,
        color="8FA4C3",
    )


def pr_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, BG)
    add_slide_title(
        slide,
        "01",
        "GitHub PR（Pull Request）",
        "向项目维护者提出代码合并申请，并在合入主分支前完成审查与讨论。",
    )

    flow_y = 1.55
    flow = [
        ("修改代码", BLUE_LIGHT, BLUE),
        ("发起 PR", BLUE_LIGHT, BLUE),
        ("审查 / 评论", PURPLE_LIGHT, PURPLE),
        ("修改迭代", ORANGE_LIGHT, ORANGE),
        ("合并主分支", GREEN_LIGHT, GREEN),
    ]
    x_positions = [0.75, 3.23, 5.71, 8.19, 10.67]
    for index, ((label, fill, color), x) in enumerate(zip(flow, x_positions)):
        add_round_rect(slide, x, flow_y, 1.9, 0.78, fill=fill, line=fill)
        add_text(
            slide,
            label,
            x,
            flow_y,
            1.9,
            0.78,
            size=15,
            color=color,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        if index < len(flow) - 1:
            add_text(
                slide,
                "→",
                x + 1.94,
                flow_y,
                0.48,
                0.78,
                size=22,
                color=MUTED,
                bold=True,
                align=PP_ALIGN.CENTER,
            )

    add_info_card(
        slide,
        0.75,
        2.75,
        3.83,
        3.7,
        "核心定义",
        "Pull Request，中文常译为“拉取请求”或“合并请求”。它不是直接改动主分支，而是提交一份可审查、可讨论的合并申请。",
        BLUE,
        BLUE_LIGHT,
    )
    add_info_card(
        slide,
        4.75,
        2.75,
        3.83,
        3.7,
        "核心作用",
        "集中展示代码差异，支持在线代码审查、逐行评论、讨论和持续修改。审核通过后，再将变更合入目标分支。",
        PURPLE,
        PURPLE_LIGHT,
    )
    add_info_card(
        slide,
        8.75,
        2.75,
        3.83,
        3.7,
        "典型场景",
        "• 为开源项目贡献代码\n• 团队内部控制合入权限\n• 在合并前运行测试和质量检查\n• 记录决策与审查过程",
        GREEN,
        GREEN_LIGHT,
    )
    add_footer(slide, 2)


def runner_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, BG)
    add_slide_title(
        slide,
        "02",
        "GitHub 后台 Agent（Actions Runner）",
        "Runner 是承载 GitHub Actions 工作流的执行程序；可使用官方云端运行器，也可自托管。",
    )

    # Central execution diagram.
    add_round_rect(slide, 4.58, 1.55, 4.18, 1.33, fill=PURPLE_LIGHT, line=PURPLE)
    add_text(
        slide,
        "ACTIONS RUNNER",
        4.58,
        1.7,
        4.18,
        0.38,
        size=12,
        color=PURPLE,
        bold=True,
        align=PP_ALIGN.CENTER,
        font_name=FONT_EN,
    )
    add_text(
        slide,
        "领取任务 · 准备环境 · 执行命令",
        4.7,
        2.13,
        3.94,
        0.36,
        size=17,
        color=TEXT,
        bold=True,
        align=PP_ALIGN.CENTER,
    )

    nodes = [
        (0.75, 1.72, 2.55, 0.9, "工作流触发", "push / PR / 定时"),
        (0.85, 3.48, 2.68, 1.22, "官方云端运行器", "即开即用\n由 GitHub 托管"),
        (9.8, 3.48, 2.68, 1.22, "自托管 Runner", "本地服务器 / 私有机器\n自行维护与控制"),
        (4.58, 4.02, 4.18, 1.18, "输出结果", "构建产物 · 测试报告 · 部署状态"),
    ]
    for index, (x, y, w, h, heading, detail) in enumerate(nodes):
        fill = WHITE if index else BLUE_LIGHT
        line = LINE if index else BLUE
        add_round_rect(slide, x, y, w, h, fill=fill, line=line)
        add_text(
            slide,
            heading,
            x + 0.15,
            y + 0.12,
            w - 0.3,
            0.34,
            size=15,
            color=TEXT,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            detail,
            x + 0.15,
            y + 0.5,
            w - 0.3,
            h - 0.58,
            size=11,
            color=MUTED,
            align=PP_ALIGN.CENTER,
        )

    add_connector(slide, 3.3, 2.17, 4.58, 2.17, color=BLUE, width=2.5)
    add_connector(slide, 6.67, 2.88, 6.67, 4.02, color=PURPLE, width=2.5)
    add_connector(slide, 4.58, 2.58, 3.54, 3.72, color=LINE, width=2)
    add_connector(slide, 8.76, 2.58, 9.8, 3.72, color=LINE, width=2)

    add_round_rect(slide, 0.75, 5.65, 11.83, 0.86, fill=NAVY, line=NAVY)
    add_text(
        slide,
        "典型任务",
        1.02,
        5.65,
        1.2,
        0.86,
        size=13,
        color="AFC2E4",
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_text(
        slide,
        "代码构建   ·   自动化测试   ·   打包制品   ·   部署发布",
        2.3,
        5.65,
        9.82,
        0.86,
        size=18,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_footer(slide, 3)


def cloud_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, BG)
    add_slide_title(
        slide,
        "03",
        "GitHub 云端集成",
        "把 GitHub 与云计算、存储、CI/CD、通知和质量平台连接起来，形成连续交付链路。",
    )

    center = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(5.15), Inches(2.34), Inches(3.02), Inches(3.02)
    )
    center.fill.solid()
    center.fill.fore_color.rgb = rgb(NAVY)
    center.line.color.rgb = rgb(NAVY)
    add_text(
        slide,
        "GitHub",
        5.15,
        2.76,
        3.02,
        0.55,
        size=27,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
        font_name=FONT_EN,
    )
    add_text(
        slide,
        "代码与事件中心",
        5.15,
        3.38,
        3.02,
        0.4,
        size=14,
        color="B7C5DC",
        align=PP_ALIGN.CENTER,
    )
    add_text(
        slide,
        "API · Webhook · App",
        5.15,
        4.06,
        3.02,
        0.36,
        size=12,
        color="8FA4C3",
        align=PP_ALIGN.CENTER,
        font_name=FONT_EN,
    )

    integrations = [
        (0.72, 1.55, 3.06, 1.08, "云端部署", "云主机 · 容器 · Serverless", BLUE, BLUE_LIGHT),
        (9.55, 1.55, 3.06, 1.08, "通知协作", "Slack · 邮件 · 工单", PURPLE, PURPLE_LIGHT),
        (0.72, 4.76, 3.06, 1.08, "质量与测试", "自动测试 · 安全扫描", ORANGE, ORANGE_LIGHT),
        (9.55, 4.76, 3.06, 1.08, "存储与制品", "镜像 · 包 · 构建产物", GREEN, GREEN_LIGHT),
    ]
    for x, y, w, h, heading, detail, color, light in integrations:
        add_round_rect(slide, x, y, w, h, fill=WHITE, line=LINE)
        add_round_rect(slide, x + 0.18, y + 0.21, 0.16, 0.66, fill=color, line=color)
        add_text(
            slide,
            heading,
            x + 0.48,
            y + 0.13,
            w - 0.66,
            0.38,
            size=16,
            color=color,
            bold=True,
        )
        add_text(
            slide,
            detail,
            x + 0.48,
            y + 0.54,
            w - 0.66,
            0.3,
            size=11,
            color=MUTED,
        )

    add_connector(slide, 3.78, 2.08, 5.35, 3.02, color=BLUE, width=2)
    add_connector(slide, 9.55, 2.08, 7.97, 3.02, color=PURPLE, width=2)
    add_connector(slide, 3.78, 5.3, 5.35, 4.65, color=ORANGE, width=2)
    add_connector(slide, 9.55, 5.3, 7.97, 4.65, color=GREEN, width=2)

    add_round_rect(slide, 4.25, 5.85, 4.83, 0.66, fill=GREEN_LIGHT, line=GREEN_LIGHT)
    add_text(
        slide,
        "提交代码后：自动构建 → 测试 → 部署 → 通知",
        4.25,
        5.85,
        4.83,
        0.66,
        size=14,
        color=GREEN,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_footer(slide, 4)


def overview_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, BG)
    add_slide_title(
        slide,
        "04",
        "三者如何协同",
        "PR 管理“是否合并”，Runner 负责“在哪里执行”，云端集成决定“连接哪些外部能力”。",
    )

    steps = [
        (0.62, "1", "提交代码", "开发者推送分支", BLUE),
        (3.14, "2", "PR 审查", "讨论、测试、批准", BLUE),
        (5.66, "3", "Runner 执行", "构建、测试、打包", PURPLE),
        (8.18, "4", "云端集成", "部署、通知、存储", GREEN),
        (10.7, "5", "交付上线", "反馈状态与结果", GREEN),
    ]
    for index, (x, number, heading, detail, color) in enumerate(steps):
        add_round_rect(slide, x, 1.58, 2.03, 1.72, fill=WHITE, line=LINE)
        number_shape = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(x + 0.72),
            Inches(1.33),
            Inches(0.58),
            Inches(0.58),
        )
        number_shape.fill.solid()
        number_shape.fill.fore_color.rgb = rgb(color)
        number_shape.line.fill.background()
        add_text(
            slide,
            number,
            x + 0.72,
            1.33,
            0.58,
            0.58,
            size=14,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            heading,
            x + 0.16,
            2.02,
            1.71,
            0.4,
            size=16,
            color=TEXT,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            detail,
            x + 0.16,
            2.49,
            1.71,
            0.48,
            size=11,
            color=MUTED,
            align=PP_ALIGN.CENTER,
        )
        if index < len(steps) - 1:
            add_text(
                slide,
                "→",
                x + 2.04,
                1.96,
                0.48,
                0.54,
                size=22,
                color=MUTED,
                bold=True,
                align=PP_ALIGN.CENTER,
            )

    add_text(slide, "快速区分", 0.68, 3.78, 2.0, 0.42, size=20, bold=True)
    comparisons = [
        (
            "PR",
            "协作机制",
            "让代码变更可审查、可讨论、可批准",
            BLUE,
            BLUE_LIGHT,
        ),
        (
            "Runner",
            "执行载体",
            "实际运行工作流中的命令与任务",
            PURPLE,
            PURPLE_LIGHT,
        ),
        (
            "云端集成",
            "连接能力",
            "把 GitHub 事件与外部云服务串联",
            GREEN,
            GREEN_LIGHT,
        ),
    ]
    for idx, (name, category, description, color, light) in enumerate(comparisons):
        y = 4.35 + idx * 0.76
        add_round_rect(slide, 0.68, y, 11.95, 0.6, fill=WHITE, line=LINE)
        add_round_rect(slide, 0.88, y + 0.12, 1.32, 0.36, fill=light, line=light)
        add_text(
            slide,
            name,
            0.88,
            y + 0.12,
            1.32,
            0.36,
            size=12,
            color=color,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            category,
            2.48,
            y + 0.08,
            1.55,
            0.44,
            size=14,
            color=TEXT,
            bold=True,
        )
        add_text(
            slide,
            description,
            4.28,
            y + 0.08,
            7.9,
            0.44,
            size=13,
            color=MUTED,
        )
    add_footer(slide, 5)


def create_presentation() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    prs.core_properties.title = "GitHub 三大核心概念"
    prs.core_properties.subject = "Pull Request、Actions Runner 与云端集成"
    prs.core_properties.author = "PPT-Producer"
    title_slide(prs)
    pr_slide(prs)
    runner_slide(prs)
    cloud_slide(prs)
    overview_slide(prs)
    prs.save(PPTX_PATH)


def text_font(size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_FILE, size)


def draw_multiline_cell(
    draw: ImageDraw.ImageDraw,
    box: tuple[int, int, int, int],
    lines: list[tuple[str, str]],
    *,
    size: int = 31,
) -> None:
    x0, y0, x1, y1 = box
    total_height = len(lines) * (size + 13)
    y = y0 + (y1 - y0 - total_height) / 2
    x = x0 + 16
    for text, color in lines:
        draw.text((x, y), text, font=text_font(size), fill=color)
        y += size + 13


def create_reference_image() -> None:
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)
    width, height = 1640, 504
    image = Image.new("RGB", (width, height), "#202224")
    draw = ImageDraw.Draw(image)
    col_x = [24, 280, 755, 1120, 1616]
    row_y = [4, 128, 301, 500]
    border = "#77808A"
    for x in col_x:
        draw.line((x, row_y[0], x, row_y[-1]), fill=border, width=2)
    for y in row_y:
        draw.line((col_x[0], y, col_x[-1], y), fill=border, width=2)

    blue = "#7FA4FF"
    white = "#F5F7FB"
    cells = [
        (0, 0, [("助理教授", blue)]),
        (1, 0, [("Assistant Professor (Asst.", white), ("Prof.)", white)]),
        (2, 0, [("Lecturer (PhD Degree", white), ("expected)", white)]),
        (3, 0, [("助教（Assistant professor /", white), ("Research Associate）", white)]),
        (0, 1, [("讲师／导师／专", blue), ("任导师", white)]),
        (1, 1, [("Lecturer / Instructor", white)]),
        (
            2,
            1,
            [
                ("Associate Lecturer /", white),
                ("Teaching Fellow /", white),
                ("Instructor", white),
            ],
        ),
        (3, 1, [("讲师（Lecturer）", white)]),
        (0, 2, [("助教／教学助理", blue)]),
        (1, 2, [("Teaching Assistant（T.A.）", white)]),
        (2, 2, [("Tutor", white)]),
        (
            3,
            2,
            [
                ("助手（Research Assistant /", white),
                ("Assistant / Associate）", white),
            ],
        ),
    ]
    for col, row, lines in cells:
        draw_multiline_cell(
            draw,
            (col_x[col], row_y[row], col_x[col + 1], row_y[row + 1]),
            lines,
        )
    image.save(REFERENCE_PATH, format="PNG", dpi=(300, 300), optimize=True)


if __name__ == "__main__":
    create_presentation()
    create_reference_image()
    print(PPTX_PATH)
    print(REFERENCE_PATH)
