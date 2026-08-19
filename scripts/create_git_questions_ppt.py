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
OUTPUT_DIR = ROOT / "Files" / "PPT-Producer"
QUESTIONS_PPTX = OUTPUT_DIR / "Git使用过程短问题.pptx"
CONCEPTS_PPTX = (
    OUTPUT_DIR
    / "Teaching-Associate教授の助理 职位"
    / "GitHub三大核心概念.pptx"
)
SCREENSHOT = Path("/tmp/166-commits-ahead.png")
FONT_FILE = "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc"

FONT_CN = "Microsoft YaHei"
FONT_EN = "Aptos"
NAVY = "16243A"
WHITE = "FFFFFF"
BG = "F5F7FB"
TEXT = "24324A"
MUTED = "6B778C"
LINE = "DCE3ED"
BLUE = "2F6BFF"
BLUE_LIGHT = "EAF0FF"
PURPLE = "7757E8"
PURPLE_LIGHT = "F0EDFF"
GREEN = "15966A"
GREEN_LIGHT = "E5F6EF"
ORANGE = "E8872E"
ORANGE_LIGHT = "FFF1E3"
RED = "D94A4A"
RED_LIGHT = "FDECEC"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def add_text(
    slide,
    value: str,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    size: int = 18,
    color: str = TEXT,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.MIDDLE,
    font_name: str = FONT_CN,
    margin: float = 0.06,
):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = shape.text_frame
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
    run.text = value
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    return shape


def add_box(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill: str = WHITE,
    line: str = LINE,
    rounded: bool = True,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    shape.line.color.rgb = rgb(line)
    shape.line.width = Pt(1)
    return shape


def set_background(slide, color: str = BG) -> None:
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(color)


def add_header(slide, number: str, title: str, subtitle: str) -> None:
    badge = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(0.62), Inches(0.42), Inches(0.54), Inches(0.54)
    )
    badge.fill.solid()
    badge.fill.fore_color.rgb = rgb(BLUE)
    badge.line.fill.background()
    add_text(
        slide,
        number,
        0.62,
        0.42,
        0.54,
        0.54,
        size=12,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_text(slide, title, 1.3, 0.3, 10.7, 0.45, size=25, bold=True)
    add_text(slide, subtitle, 1.32, 0.78, 10.8, 0.3, size=11, color=MUTED)
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.65), Inches(1.16), Inches(12.0), Inches(0.01)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = rgb(LINE)
    line.line.fill.background()


def add_footer(slide, page: int) -> None:
    add_text(
        slide,
        "Git 使用过程短问题",
        0.66,
        7.13,
        3.0,
        0.2,
        size=9,
        color=MUTED,
    )
    add_text(
        slide,
        f"{page:02d}",
        12.18,
        7.12,
        0.48,
        0.2,
        size=9,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )


def add_qa_card(
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
    add_box(slide, x, y, w, h)
    add_box(slide, x + 0.22, y + 0.2, 0.9, 0.38, fill=light, line=light)
    add_text(
        slide,
        label,
        x + 0.22,
        y + 0.2,
        0.9,
        0.38,
        size=11,
        color=color,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_text(
        slide,
        body,
        x + 0.24,
        y + 0.72,
        w - 0.48,
        h - 0.9,
        size=16,
        valign=MSO_ANCHOR.TOP,
    )


def create_ahead_screenshot() -> None:
    image = Image.new("RGB", (996, 114), "#F6F8FA")
    draw = ImageDraw.Draw(image)
    draw.line((0, 112, 996, 112), fill="#D0D7DE", width=2)
    regular = ImageFont.truetype(FONT_FILE, 29)
    blue_font = ImageFont.truetype(FONT_FILE, 29)
    x, y = 28, 37
    segments = [
        ("This branch is ", "#24292F", regular),
        ("166 commits ahead of", "#0969DA", blue_font),
        (" ", "#24292F", regular),
        ("LeonYanghaha/pua-books:master", "#57606A", regular),
        (".", "#24292F", regular),
    ]
    for value, color, selected_font in segments:
        bounds = draw.textbbox((0, 0), value, font=selected_font)
        if color in {"#0969DA", "#57606A"}:
            padding = 4
            draw.rounded_rectangle(
                (x - padding, y - 3, x + bounds[2] + padding, y + 36),
                radius=4,
                fill="#EAF3FF" if color == "#0969DA" else "#EAEEF2",
            )
        draw.text((x, y), value, font=selected_font, fill=color)
        x += bounds[2] - bounds[0]
    image.save(SCREENSHOT, format="PNG", dpi=(200, 200), optimize=True)


def title_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, NAVY)
    add_box(slide, 0.72, 0.62, 1.65, 0.36, fill="223753", line="223753")
    add_text(
        slide,
        "GIT · QUICK HELP",
        0.72,
        0.62,
        1.65,
        0.36,
        size=10,
        color="ADC0DD",
        bold=True,
        align=PP_ALIGN.CENTER,
        font_name=FONT_EN,
    )
    add_text(
        slide,
        "Git 使用过程短问题",
        0.72,
        1.56,
        9.1,
        0.72,
        size=36,
        color=WHITE,
        bold=True,
    )
    add_text(
        slide,
        "记录在 GitHub 使用过程中遇到的常见疑问与处理方法",
        0.75,
        2.35,
        9.8,
        0.42,
        size=18,
        color="B8C6DC",
    )
    topics = [
        (0.75, "01", "PR 与分支", BLUE),
        (3.82, "02", "二进制文件", PURPLE),
        (6.89, "03", "替换 PPTX", ORANGE),
        (9.96, "04", "提交差异", GREEN),
    ]
    for x, number, label, color in topics:
        add_box(slide, x, 3.52, 2.62, 1.62, fill="20324C", line="3B4E69")
        add_box(slide, x + 0.2, 3.75, 0.5, 0.5, fill=color, line=color)
        add_text(
            slide,
            number,
            x + 0.2,
            3.75,
            0.5,
            0.5,
            size=12,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            label,
            x + 0.2,
            4.45,
            2.2,
            0.34,
            size=16,
            color=WHITE,
            bold=True,
        )
    add_text(
        slide,
        "个人 GitHub 操作备忘",
        0.76,
        6.83,
        4.5,
        0.25,
        size=11,
        color="8EA2C1",
    )


def pr_branch_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide)
    add_header(
        slide,
        "01",
        "PR 不合并，是否还能继续修改？",
        "PR 是合并申请；真正承载文件变更的是来源分支。",
    )
    add_qa_card(
        slide,
        0.72,
        1.52,
        3.72,
        2.05,
        "回答",
        "可以。只要来源分支仍存在，就能继续修改、提交和推送；现有 PR 会自动同步这些更新。",
        GREEN,
        GREEN_LIGHT,
    )
    add_qa_card(
        slide,
        4.8,
        1.52,
        3.72,
        2.05,
        "注意",
        "关闭浏览器不等于关闭 PR。只有主动点击“Close pull request”，PR 才会被关闭。",
        ORANGE,
        ORANGE_LIGHT,
    )
    add_qa_card(
        slide,
        8.88,
        1.52,
        3.72,
        2.05,
        "建议",
        "修改期间保留为 Draft；全部定稿后再合并。个人存档可使用 Squash and merge。",
        BLUE,
        BLUE_LIGHT,
    )

    labels = [
        (0.85, "功能分支", "持续修改文件", BLUE, BLUE_LIGHT),
        (4.1, "Pull Request", "自动显示更新", PURPLE, PURPLE_LIGHT),
        (7.35, "审核 / 定稿", "确认最终内容", ORANGE, ORANGE_LIGHT),
        (10.6, "main", "合并后正式保存", GREEN, GREEN_LIGHT),
    ]
    for index, (x, heading, detail, color, light) in enumerate(labels):
        add_box(slide, x, 4.37, 2.05, 1.25, fill=light, line=light)
        add_text(
            slide,
            heading,
            x,
            4.56,
            2.05,
            0.34,
            size=15,
            color=color,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            detail,
            x,
            5.0,
            2.05,
            0.28,
            size=11,
            color=MUTED,
            align=PP_ALIGN.CENTER,
        )
        if index < len(labels) - 1:
            add_text(
                slide,
                "→",
                x + 2.18,
                4.67,
                0.72,
                0.4,
                size=22,
                color=MUTED,
                bold=True,
                align=PP_ALIGN.CENTER,
            )
    add_footer(slide, 2)


def binary_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide)
    add_header(
        slide,
        "02",
        "“Binary file not shown”是什么意思？",
        "GitHub 可以保存二进制文件，但无法像文本代码一样逐行展示差异。",
    )
    add_box(slide, 0.78, 1.52, 4.05, 4.78, fill=NAVY, line=NAVY)
    add_text(
        slide,
        "Binary file\nnot shown.",
        1.17,
        2.13,
        3.26,
        1.2,
        size=29,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
        font_name=FONT_EN,
    )
    add_text(
        slide,
        "无法显示内容差异\n≠\n文件不存在或损坏",
        1.2,
        3.62,
        3.2,
        1.38,
        size=17,
        color="B8C6DC",
        align=PP_ALIGN.CENTER,
    )

    add_qa_card(
        slide,
        5.2,
        1.52,
        3.45,
        2.2,
        "常见类型",
        "PPTX、DOCX、PNG、JPG、ZIP 等文件通常属于二进制文件。",
        PURPLE,
        PURPLE_LIGHT,
    )
    add_qa_card(
        slide,
        8.98,
        1.52,
        3.45,
        2.2,
        "仍可操作",
        "仍然可以下载、删除、替换、提交和合并，只是 GitHub 不提供逐行预览。",
        GREEN,
        GREEN_LIGHT,
    )
    add_qa_card(
        slide,
        5.2,
        4.05,
        7.23,
        2.25,
        "PPT 修改方式",
        "先下载并在 WPS 中编辑，再切换到正确分支和文件夹，使用 Add file → Upload files 上传同名文件进行替换。",
        BLUE,
        BLUE_LIGHT,
    )
    add_footer(slide, 3)


def replace_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide)
    add_header(
        slide,
        "03",
        "如何替换已修改的 PPTX？",
        "关键是选择正确分支、保持路径和文件名一致，并提交到 PR 的来源分支。",
    )
    steps = [
        ("1", "选择分支", "切换到 PR 的来源分支", BLUE, BLUE_LIGHT),
        ("2", "进入目录", "打开 PPT 所在文件夹", PURPLE, PURPLE_LIGHT),
        ("3", "上传同名文件", "Add file → Upload files", ORANGE, ORANGE_LIGHT),
        ("4", "提交更改", "Commit 到当前分支", GREEN, GREEN_LIGHT),
    ]
    x_positions = [0.7, 3.84, 6.98, 10.12]
    for index, ((number, heading, detail, color, light), x) in enumerate(
        zip(steps, x_positions)
    ):
        add_box(slide, x, 1.65, 2.5, 2.35, fill=WHITE, line=LINE)
        badge = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            Inches(x + 0.88),
            Inches(1.42),
            Inches(0.74),
            Inches(0.74),
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = rgb(color)
        badge.line.fill.background()
        add_text(
            slide,
            number,
            x + 0.88,
            1.42,
            0.74,
            0.74,
            size=16,
            color=WHITE,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            heading,
            x + 0.2,
            2.34,
            2.1,
            0.4,
            size=17,
            color=color,
            bold=True,
            align=PP_ALIGN.CENTER,
        )
        add_text(
            slide,
            detail,
            x + 0.22,
            2.91,
            2.06,
            0.62,
            size=12,
            color=MUTED,
            align=PP_ALIGN.CENTER,
        )
        if index < 3:
            add_text(
                slide,
                "→",
                x + 2.54,
                2.39,
                0.58,
                0.45,
                size=23,
                color=MUTED,
                bold=True,
                align=PP_ALIGN.CENTER,
            )
    add_box(slide, 0.7, 4.52, 11.92, 1.35, fill=RED_LIGHT, line=RED_LIGHT)
    add_text(
        slide,
        "避免这三个错误",
        0.98,
        4.75,
        1.72,
        0.35,
        size=14,
        color=RED,
        bold=True,
    )
    add_text(
        slide,
        "① 提交到 main　　② 文件名出现“(1)”　　③ 上传到错误文件夹",
        2.8,
        4.63,
        8.95,
        0.62,
        size=16,
        color=TEXT,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_text(
        slide,
        "上传成功后，现有 PR 会自动增加一个新提交并显示最新文件。",
        1.0,
        5.38,
        10.6,
        0.3,
        size=12,
        color=MUTED,
        align=PP_ALIGN.CENTER,
    )
    add_footer(slide, 4)


def branches_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide)
    add_header(
        slide,
        "04",
        "Branches 是不是类似硬盘分区？",
        "分支不是物理分区，而是同一仓库中指向不同提交历史的平行版本线。",
    )

    add_box(slide, 0.78, 1.62, 3.15, 4.65, fill=NAVY, line=NAVY)
    add_text(
        slide,
        "仓库",
        1.05,
        1.95,
        2.6,
        0.44,
        size=23,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_text(
        slide,
        "同一套文件与历史",
        1.05,
        2.5,
        2.6,
        0.34,
        size=13,
        color="B8C6DC",
        align=PP_ALIGN.CENTER,
    )
    add_box(slide, 1.2, 3.26, 2.32, 0.74, fill=BLUE, line=BLUE)
    add_text(
        slide,
        "main",
        1.2,
        3.26,
        2.32,
        0.74,
        size=18,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
        font_name=FONT_EN,
    )
    add_box(slide, 1.2, 4.47, 2.32, 0.98, fill=PURPLE, line=PURPLE)
    add_text(
        slide,
        "feature branch\n含未合并的 PPT",
        1.2,
        4.47,
        2.32,
        0.98,
        size=15,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )

    add_qa_card(
        slide,
        4.38,
        1.62,
        3.75,
        2.15,
        "为什么 main 看不到？",
        "PPT 只存在于功能分支，尚未合并到 main。切换分支后，仓库文件视图会随之变化。",
        BLUE,
        BLUE_LIGHT,
    )
    add_qa_card(
        slide,
        8.52,
        1.62,
        3.75,
        2.15,
        "合并后会怎样？",
        "PR 合并后，功能分支的最终内容进入 main，主分支页面才会显示这些文件。",
        GREEN,
        GREEN_LIGHT,
    )
    add_qa_card(
        slide,
        4.38,
        4.16,
        7.89,
        2.11,
        "简单理解",
        "分支更像“同一文档的平行版本”，不是把硬盘切成多个空间。每个分支都记录自己当前指向的版本。",
        PURPLE,
        PURPLE_LIGHT,
    )
    add_footer(slide, 5)


def ahead_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide)
    add_header(
        slide,
        "05",
        "“166 commits ahead”是什么意思？",
        "ahead 描述的是提交历史差异，不是文件数量，也不是错误提示。",
    )
    slide.shapes.add_picture(
        str(SCREENSHOT), Inches(0.75), Inches(1.48), width=Inches(11.85)
    )
    add_qa_card(
        slide,
        0.75,
        2.98,
        3.72,
        2.65,
        "中文含义",
        "当前分支比 LeonYanghaha/pua-books 的 master 分支领先 166 个提交。",
        BLUE,
        BLUE_LIGHT,
    )
    add_qa_card(
        slide,
        4.8,
        2.98,
        3.72,
        2.65,
        "代表什么",
        "这些提交存在于当前分支，但目标 master 尚未包含。它可能来自持续开发、分叉或历史未同步。",
        PURPLE,
        PURPLE_LIGHT,
    )
    add_qa_card(
        slide,
        8.85,
        2.98,
        3.72,
        2.65,
        "是否需要处理",
        "如果这是预期的开发差异，无需担心；准备合并前，应确认这 166 个提交确实都应进入目标分支。",
        GREEN,
        GREEN_LIGHT,
    )
    add_box(slide, 0.75, 5.96, 11.82, 0.57, fill=ORANGE_LIGHT, line=ORANGE_LIGHT)
    add_text(
        slide,
        "提醒：166 commits ≠ 166 个文件。一个提交可以修改零个、一个或多个文件。",
        0.95,
        5.96,
        11.42,
        0.57,
        size=14,
        color=ORANGE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_footer(slide, 6)


def create_questions_presentation() -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    prs.core_properties.title = "Git 使用过程短问题"
    prs.core_properties.subject = "GitHub 常见操作疑问与处理方法"
    prs.core_properties.author = "PPT-Producer"
    title_slide(prs)
    pr_branch_slide(prs)
    binary_slide(prs)
    replace_slide(prs)
    branches_slide(prs)
    ahead_slide(prs)
    prs.save(QUESTIONS_PPTX)


def append_ahead_slide_to_concepts() -> None:
    prs = Presentation(CONCEPTS_PPTX)
    if len(prs.slides) != 6:
        raise RuntimeError(
            f"Expected current six-slide concepts deck, found {len(prs.slides)} slides"
        )
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_background(slide, BG)
    add_text(
        slide,
        "“166 commits ahead”是什么意思？",
        0.72,
        0.38,
        7.7,
        0.52,
        size=26,
        color=TEXT,
        bold=True,
    )
    slide.shapes.add_picture(
        str(SCREENSHOT), Inches(0.74), Inches(1.18), width=Inches(11.85)
    )
    add_box(slide, 0.74, 2.25, 5.7, 3.7, fill=WHITE, line=LINE)
    add_text(
        slide,
        "简短回答",
        1.02,
        2.53,
        1.45,
        0.38,
        size=15,
        color=BLUE,
        bold=True,
    )
    add_text(
        slide,
        "当前分支中有 166 个提交尚未包含在目标仓库的 master 分支中。这不是报错，也不是 166 个文件；合并前应确认这些提交都是预期内容。",
        1.02,
        3.12,
        5.15,
        2.25,
        size=18,
        color=TEXT,
        valign=MSO_ANCHOR.TOP,
    )

    add_box(slide, 6.8, 2.25, 5.78, 1.64, fill=BLUE_LIGHT, line=BLUE_LIGHT)
    add_text(
        slide,
        "英文 OCR",
        7.06,
        2.47,
        1.4,
        0.32,
        size=15,
        color=BLUE,
        bold=True,
    )
    add_text(
        slide,
        "This branch is 166 commits ahead of\nLeonYanghaha/pua-books:master.",
        7.06,
        2.9,
        5.1,
        0.7,
        size=15,
        color=TEXT,
        font_name=FONT_EN,
        valign=MSO_ANCHOR.TOP,
    )

    add_box(slide, 6.8, 4.3, 5.78, 1.64, fill=GREEN_LIGHT, line=GREEN_LIGHT)
    add_text(
        slide,
        "中文翻译",
        7.06,
        4.52,
        1.4,
        0.32,
        size=15,
        color=GREEN,
        bold=True,
    )
    add_text(
        slide,
        "此分支比 LeonYanghaha/pua-books 的\nmaster 分支领先 166 个提交。",
        7.06,
        4.95,
        5.1,
        0.7,
        size=15,
        color=TEXT,
        valign=MSO_ANCHOR.TOP,
    )
    add_text(
        slide,
        "注：ahead 只比较提交历史是否已被目标分支包含。",
        0.78,
        6.56,
        8.5,
        0.3,
        size=11,
        color=MUTED,
    )

    with NamedTemporaryFile(
        suffix=".pptx", dir=CONCEPTS_PPTX.parent, delete=False
    ) as temp_file:
        temp_path = Path(temp_file.name)
    try:
        prs.save(temp_path)
        temp_path.replace(CONCEPTS_PPTX)
    finally:
        temp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    create_ahead_screenshot()
    create_questions_presentation()
    append_ahead_slide_to_concepts()
    print(QUESTIONS_PPTX)
    print(CONCEPTS_PPTX)
