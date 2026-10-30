"""Pure test data for Word/AxMath typography verification.

The 14 cases cover four pages and reuse seven expressions. A verification runner
may reuse already verified native OLE objects across contexts to reduce conversion
time. Production conversion must not treat that test-only reuse as a general cache.
Expected tokens are lightweight source-retention checks, not proofs of equivalent
mathematics or correct typography; the rendered pages still need visual inspection.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal

Layout = Literal["inline", "display", "numbered"]
Language = Literal["zh", "en", "mixed"]
LineSpacing = Literal["single", "at_least", "multiple"]


@dataclass(frozen=True)
class LayoutCase:
    case_id: str
    page: int
    title: str
    latex: str
    layout: Layout
    before: str
    after: str
    expected_tokens: tuple[str, ...]
    language: Language = "mixed"
    font_size_pt: float = 12.0
    line_spacing: LineSpacing = "at_least"
    line_spacing_value: float = 18.0


POWER = r"x_i^2+\alpha_i"
FRACTION = r"\frac{a+b}{c+d}"
ROOT = r"\sqrt{x^2+y^2}"
SUM = r"\sum_{i=1}^{n}x_i=n\bar{x}"
GAUSSIAN = r"\int_{-\infty}^{+\infty}e^{-x^2}\,dx=\sqrt{\pi}"
# Columns sum to one, so a column probability vector evolves by p_(t+1) = P p_t.
MARKOV = r"\mathbf{p}_{t+1}=\begin{pmatrix}0.8&0.3\\0.2&0.7\end{pmatrix}\mathbf{p}_t"
PIECEWISE = r"f(x)=\begin{cases}x^2,&x\geq0\\-x,&x<0\end{cases}"

PAGE_TITLES = {
    1: "中文正文中的行内公式",
    2: "Inline equations in English text",
    3: "独立居中公式",
    4: "带自动编号的公式",
}

PAGE_INTRODUCTIONS = {
    1: "检查中文与公式的基线、标点间距及上下行间距。三个段落分别包含上下标、分式和根式。",
    2: (
        "Inspect the baseline, punctuation spacing and line height when editable equations "
        "appear within English paragraphs. The expressions match the Chinese examples."
    ),
    3: "检查求和、积分、矩阵和分段函数是否独立居中，公式上下是否留有足够空间。",
    4: (
        "检查公式居中、编号靠右以及两者互不挤占。编号由 Word 序列域生成，"
        "应在更新公式后仍连续，保存并重新打开后仍可更新。"
    ),
}

CASES: tuple[LayoutCase, ...] = (
    LayoutCase(
        case_id="zh_inline_power",
        page=1,
        title="中文与上下标",
        latex=POWER,
        layout="inline",
        language="zh",
        before="对于第 i 个观测值，将调整后的量记为",
        after=(
            "，其中第二项表示该观测值的修正量。公式前后的中文应处于自然的阅读基线上，"
            "逗号紧随公式，上标和下标均应完整显示，换行后的文字不得与公式发生碰撞。"
        ),
        expected_tokens=("x", "i", "2", r"\alpha"),
    ),
    LayoutCase(
        case_id="zh_inline_fraction",
        page=1,
        title="中文与分式",
        latex=FRACTION,
        layout="inline",
        language="zh",
        before="当分母非零时，这两个合计量的比值写为",
        after=(
            "。分子与分母都应完整显示，中文句号应与邻近文字保持一致的基线。"
            "本段保留足够的后续文字，以观察含分式的一行和普通文字行之间的间距是否自然。"
        ),
        expected_tokens=(r"\frac", "a", "b", "c", "d"),
    ),
    LayoutCase(
        case_id="zh_inline_root",
        page=1,
        title="中文与根式",
        latex=ROOT,
        layout="inline",
        language="zh",
        before="在二维坐标系中，点到原点的距离为",
        after=(
            "，这一表达式同时包含根号和上标。应检查根号横线未被裁切，"
            "公式与中文之间没有异常空白，且同一段落中的其他文字未因公式插入而改变字号。"
        ),
        expected_tokens=(r"\sqrt", "x", "y", "2"),
    ),
    LayoutCase(
        case_id="en_inline_power",
        page=2,
        title="English and scripts",
        latex=POWER,
        layout="inline",
        language="en",
        before="For observation i, define the adjusted quantity as ",
        after=(
            ", where the second term is an observation-specific correction. "
            "The superscript and subscript should remain legible, while the surrounding "
            "letters and the comma share a natural reading baseline across wrapped lines."
        ),
        expected_tokens=("x", "i", "2", r"\alpha"),
    ),
    LayoutCase(
        case_id="en_inline_fraction",
        page=2,
        title="English and a fraction",
        latex=FRACTION,
        layout="inline",
        language="en",
        before="When the denominator is nonzero, the ratio of the two totals is ",
        after=(
            ". Both parts of the fraction must remain visible. This longer paragraph "
            "also makes it possible to inspect the spacing between the line containing "
            "the equation and the adjacent lines of ordinary English text."
        ),
        expected_tokens=(r"\frac", "a", "b", "c", "d"),
    ),
    LayoutCase(
        case_id="en_inline_root",
        page=2,
        title="English and a square root",
        latex=ROOT,
        layout="inline",
        language="en",
        before="The Euclidean distance from the point to the origin is ",
        after=(
            ", measured in the same units as the coordinates. Inspect the top of the "
            "radical, the position of the following comma, and the whitespace around "
            "the expression without changing the surrounding text size."
        ),
        expected_tokens=(r"\sqrt", "x", "y", "2"),
    ),
    LayoutCase(
        case_id="display_sum",
        page=3,
        title="求和与上下限",
        latex=SUM,
        layout="display",
        before="对于 n 个观测值，总和与样本均值满足以下关系。",
        after="The summation limits and the bar over the mean must remain visible.",
        expected_tokens=(r"\sum", "i", "n", r"\bar"),
    ),
    LayoutCase(
        case_id="display_gaussian",
        page=3,
        title="高斯积分",
        latex=GAUSSIAN,
        layout="display",
        before="高斯积分提供了含积分上下限、指数和根号的排版样例。",
        after="Check the integral limits, exponent and radical without clipping.",
        expected_tokens=(r"\int", r"\infty", "e", r"\sqrt", r"\pi"),
    ),
    LayoutCase(
        case_id="display_markov",
        page=3,
        title="矩阵与向量",
        latex=MARKOV,
        layout="display",
        before="列概率向量按列随机转移矩阵进行一步更新。",
        after="The two matrix rows should align within complete parentheses.",
        expected_tokens=("p", "t", "0.8", "0.3", "0.2", "0.7"),
    ),
    LayoutCase(
        case_id="display_piecewise",
        page=3,
        title="分段函数",
        latex=PIECEWISE,
        layout="display",
        before="分段函数用于检查大括号、分行以及条件对齐。",
        after="Both branches and their conditions should be readable as separate rows.",
        expected_tokens=("f", "x", "2", r"\geq", "0"),
    ),
    LayoutCase(
        case_id="numbered_gaussian",
        page=4,
        title="带编号的积分",
        latex=GAUSSIAN,
        layout="numbered",
        before="将同一积分放入带编号的独立公式段落。",
        after="The equation remains centered while its number is aligned to the right.",
        expected_tokens=(r"\int", r"\infty", "e", r"\sqrt", r"\pi"),
    ),
    LayoutCase(
        case_id="numbered_markov",
        page=4,
        title="带编号的矩阵",
        latex=MARKOV,
        layout="numbered",
        before="矩阵公式需要容纳两行内容，编号不应进入矩阵内部。",
        after="The number must remain separate from the tall, editable matrix object.",
        expected_tokens=("p", "t", "0.8", "0.3", "0.2", "0.7"),
    ),
    LayoutCase(
        case_id="numbered_piecewise",
        page=4,
        title="带编号的分段函数",
        latex=PIECEWISE,
        layout="numbered",
        before="分段函数与其编号应作为同一个排版单元保留。",
        after="Neither the branches nor the number should be clipped or overlap.",
        expected_tokens=("f", "x", "2", r"\geq", "0"),
    ),
    LayoutCase(
        case_id="numbered_sum",
        page=4,
        title="连续编号与公式更新",
        latex=SUM,
        layout="numbered",
        before="最后一个公式用于检查编号是否连续，以及替换公式后编号是否保留。",
        after="Update the formula object and refresh fields without adding a duplicate number.",
        expected_tokens=(r"\sum", "i", "n", r"\bar"),
    ),
)

CASE_BY_ID = {case.case_id: case for case in CASES}
