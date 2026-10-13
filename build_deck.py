from pathlib import Path
from pptx import Presentation 
from analysis import get_kpis 
from pptx.dml.color import RGBColor
import json



# ============================================================
# 1. FILE PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_PATH = BASE_DIR /  "novaretail_template.pptx"
CHARTS_DIR = BASE_DIR / "charts"
ICONS_DIR = BASE_DIR / "novaretail_icon_repo" 
OUTPUT_DIR = BASE_DIR / "output"

OUTPUT_PATH = OUTPUT_DIR / "NovaRetail_QBR.pptx"


# Chart file paths
REVENUE_CHART = CHARTS_DIR / "revenue.png"
CUSTOMER_CHART = CHARTS_DIR / "customer.png"
RISK_CHART = CHARTS_DIR / "stockout_rate_pct.png"


# Icon file paths
# Change these filenames if your extracted icons have different names.
ICON_1 = ICONS_DIR / "globe_region.png"
ICON_2 = ICONS_DIR / "growth_chart.png"
ICON_3 = ICONS_DIR / "customer_star.png"
ICON_4 = ICONS_DIR / "risk_alert.png"
ICON_5 = ICONS_DIR / "lightbulb_action.png"


# ============================================================
# 2. HELPER FUNCTIONS
# ============================================================

def find_shape(slide, shape_name):
    """
    Find and return a shape using its exact shape name.
    """

    for shape in slide.shapes:
        if shape.name == shape_name:
            return shape

    available_shapes = [shape.name for shape in slide.shapes]

    raise KeyError(
        f"Shape '{shape_name}' was not found.\n"
        f"Available shapes: {available_shapes}"
    )


def set_text(slide, shape_name, text, color=None):

    shape = find_shape(slide, shape_name)

    if not shape.has_text_frame:
        raise TypeError(
            f"Shape '{shape_name}' does not contain a text frame."
        )

    shape.text_frame.clear()

    paragraph = shape.text_frame.paragraphs[0]
    run = paragraph.add_run()

    run.text = str(text)

    if color:
        run.font.color.rgb = color


def set_bullet_text(slide, shape_name, bullets,color=None):
    """
    Add bullet-style text to a named text shape.

    The bullet symbol is added directly because some templates
    do not preserve PowerPoint bullet formatting consistently.
    """

    shape = find_shape(slide, shape_name)

    if not shape.has_text_frame:
        raise TypeError(
            f"Shape '{shape_name}' does not contain a text frame."
        )

    text_frame = shape.text_frame
    text_frame.clear()

    for index, bullet in enumerate(bullets):
        if index == 0:
            paragraph = text_frame.paragraphs[0]
        else:
            paragraph = text_frame.add_paragraph()
            
        
        run = paragraph.add_run()
        run.text = f"• {bullet}"

        if color:
            run.font.color.rgb = color
        paragraph.level = 0
        paragraph.space_after = 0


def remove_shape(shape):
    """
    Remove a PowerPoint shape from a slide.
    """

    shape_element = shape._element
    shape_element.getparent().remove(shape_element)


def remove_shape_by_name(slide, shape_name, required=False):
    """
    Remove a named shape.

    If required=False, the script continues when the shape
    does not exist.
    """

    for shape in slide.shapes:
        if shape.name == shape_name:
            remove_shape(shape)
            return True

    if required:
        raise KeyError(f"Cannot remove missing shape '{shape_name}'.")

    return False


def replace_shape_with_picture(slide, shape_name, image_path):
    """
    Replace a named placeholder shape with an image.

    The image is placed at the exact location and size
    of the original shape.
    """

    image_path = Path(image_path)

    if not image_path.exists():
        raise FileNotFoundError(
            f"Image file was not found: {image_path}"
        )

    placeholder = find_shape(slide, shape_name)

    left = placeholder.left
    top = placeholder.top
    width = placeholder.width
    height = placeholder.height

    remove_shape(placeholder)

    picture = slide.shapes.add_picture(
        str(image_path),
        left,
        top,
        width=width,
        height=height,
    )

    return picture


def insert_icon_if_available(slide, shape_name, icon_path):
    """
    Replace an icon placeholder only when the icon file exists.

    This lets you test the presentation even if icons have not
    yet been added to the assets/icons folder.
    """

    icon_path = Path(icon_path)

    if icon_path.exists():
        replace_shape_with_picture(
            slide=slide,
            shape_name=shape_name,
            image_path=icon_path,
        )

        print(f"Inserted icon: {icon_path.name}")
        return True

    print(
        f"Warning: icon not found: {icon_path}\n"
        f"The placeholder '{shape_name}' was left unchanged."
    )

    return False

def load_deck_content():

    with open(
        "output/deck_content.json",
        "r",
        encoding="utf-8"
    ) as f:

        return json.load(f)


# ============================================================
# 3. VALIDATION FUNCTIONS
# ============================================================

def validate_input_files():
    """
    Check the required template and chart files.
    """

    required_files = [
        TEMPLATE_PATH,
        REVENUE_CHART,
        CUSTOMER_CHART,
        RISK_CHART,
    ]

    missing_files = [
        str(file_path)
        for file_path in required_files
        if not file_path.exists()
    ]

    if missing_files:
        missing_text = "\n".join(missing_files)

        raise FileNotFoundError(
            "The following required files are missing:\n"
            f"{missing_text}"
        )


def validate_slide_count(prs):
    """
    Confirm that the template contains exactly five slides.
    """

    if len(prs.slides) != 5:
        raise ValueError(
            "The PowerPoint template must contain exactly "
            f"5 slides. It currently contains {len(prs.slides)}."
        )


def validate_required_shapes(prs):
    """
    Confirm that the required named shapes exist before editing.
    """

    required_shapes = {
        1: [
            "deck_title",
            "subtitle_1",
            "kpi1_value",
            "kpi1_label",
            "kpi2_value",
            "kpi2_label",
            "kpi3_value",
            "kpi3_label",
            "summary_text",
            "icon_1",
        ],
        2: [
            "title_2",
            "chart_performance",
            "insights_performance",
            "icon_2",
        ],
        3: [
            "title_3",
            "chart_customer",
            "insights_customer",
            "icon_3",
        ],
        4: [
            "title_4",
            "chart_risk",
            "insights_risk",
            "icon_4",
        ],
        5: [
            "title_5",
            "recommendations_text",
            "icon_5",
            "closing_statement",
        ],
    }

    problems = []

    for slide_number, shape_names in required_shapes.items():
        slide = prs.slides[slide_number - 1]
        available = [shape.name for shape in slide.shapes]

        for shape_name in shape_names:
            if shape_name not in available:
                problems.append(
                    f"Slide {slide_number}: missing '{shape_name}'"
                )

    if problems:
        raise ValueError(
            "Required PowerPoint shapes are missing:\n"
            + "\n".join(problems)
        )


def validate_final_deck(prs):
    """
    Check that obvious placeholder instructions are not left
    in the final deck.
    """

    forbidden_phrases = [
        "insert chart",
        "insert icon",
        "choose icon",
        "pick icon",
        "replace this",
        "add insights",
        "enter text",
        "[value]",
        "[fill",
    ]

    problems = []

    for slide_number, slide in enumerate(prs.slides, start=1):

        for shape in slide.shapes:

            if not shape.has_text_frame:
                continue

            shape_text = shape.text.lower()

            for phrase in forbidden_phrases:
                if phrase in shape_text:
                    problems.append(
                        f"Slide {slide_number}, "
                        f"shape '{shape.name}': "
                        f"found placeholder phrase '{phrase}'"
                    )

    if problems:
        print("\nWarning: possible placeholder text remains:")

        for problem in problems:
            print(f"- {problem}")


# ============================================================
# 4. SLIDE CONTENT
# ============================================================
content = load_deck_content()
WHITE = RGBColor(255, 255, 255)
def fill_slide_1(prs):
    """
    Fill the executive summary slide.
    """

    slide = prs.slides[0]

    
    set_text(
        slide,
        "subtitle_1",
        "Review Period: Q2 2026 Board Review",
        WHITE
    )

    # KPI 1
    kpis = get_kpis()
    set_text(
        slide,
        "kpi1_value",
        kpis["revenue"],
        WHITE
    )

    set_text(
        slide,
        "kpi1_label",
        "Total Revenue",
        WHITE
    )

    # KPI 2
    set_text(
        slide,
        "kpi2_value",
        kpis["growth"],
        WHITE
    )

    set_text(
        slide,
        "kpi2_label",
        "QoQ Revenue Growth",
        WHITE
    )

    # KPI 3
    set_text(
        slide,
        "kpi3_value",
        kpis["nps"],
        WHITE
    )

    set_text(
        slide,
        "kpi3_label",
        "Average Regional NPS",
        WHITE
    )


    set_text(
        slide,
        "summary_text",
        content["summary"],
        WHITE
    )

    # Keep the existing summary_label, footer and page number.

    icon_inserted = insert_icon_if_available(
        slide,
        "icon_1",
        ICON_1,
    )

    if icon_inserted:
        remove_shape_by_name(
            slide,
            "icon_1_label",
            required=False,
        )


def fill_slide_2(prs):
    """
    Fill the regional performance slide.
    """

    slide = prs.slides[1]
    

    set_text(
        slide,
        "title_2",
         content["slide2_headline"]
    )

    

    set_bullet_text(
        slide,
        "insights_performance",
        content["performance_bullets"]
    )

    replace_shape_with_picture(
        slide,
        "chart_performance",
        REVENUE_CHART,
    )

    remove_shape_by_name(
        slide,
        "chart_performance_label",
        required=False,
    )

    icon_inserted = insert_icon_if_available(
        slide,
        "icon_2",
        ICON_2,
    )

    if icon_inserted:
        remove_shape_by_name(
            slide,
            "icon_2_label",
            required=False,
        )


def fill_slide_3(prs):
    """
    Fill the customer-experience slide.
    """

    slide = prs.slides[2]
    

    set_text(
        slide,
        "title_3",
         content["slide3_headline"]
    )

    

    set_bullet_text(
        slide,
        "insights_customer",
        content["customer_bullets"]
    )

    replace_shape_with_picture(
        slide,
        "chart_customer",
        CUSTOMER_CHART,
    )

    remove_shape_by_name(
        slide,
        "chart_customer_label",
        required=False,
    )

    icon_inserted = insert_icon_if_available(
        slide,
        "icon_3",
        ICON_3,
    )

    if icon_inserted:
        remove_shape_by_name(
            slide,
            "icon_3_label",
            required=False,
        )


def fill_slide_4(prs):
    """
    Fill the operational-risk slide.
    """

    slide = prs.slides[3]
    

    set_text(
        slide,
        "title_4",
        content["slide4_headline"]
    )

    
    set_bullet_text(
        slide,
        "insights_risk",
        content["risk_bullets"]
    )

    replace_shape_with_picture(
        slide,
        "chart_risk",
        RISK_CHART,
    )

    remove_shape_by_name(
        slide,
        "chart_risk_label",
        required=False,
    )

    icon_inserted = insert_icon_if_available(
        slide,
        "icon_4",
        ICON_4,
    )

    if icon_inserted:
        remove_shape_by_name(
            slide,
            "icon_4_label",
            required=False,
        )


def fill_slide_5(prs):
    """
    Fill the recommendations slide.
    """

    slide = prs.slides[4]

    

    
    set_bullet_text(
        slide,
        "recommendations_text",
        content["recommendations"],
        WHITE
    )

    

    set_text(
        slide,
        "closing_statement",
        content["closing_statement"],
        WHITE
    )

    icon_inserted = insert_icon_if_available(
        slide,
        "icon_5",
        ICON_5,
    )

    if icon_inserted:
        remove_shape_by_name(
            slide,
            "icon_5_label",
            required=False,
        )


# ============================================================
# 5. BUILD COMPLETE DECK
# ============================================================

def build_deck():
    """
    Open the fixed template, populate all five slides and save
    the final presentation.
    """

    print("Starting NovaRetail PowerPoint generation...")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    validate_input_files()

    print(f"Opening template: {TEMPLATE_PATH}")

    prs = Presentation(str(TEMPLATE_PATH))

    validate_slide_count(prs)
    validate_required_shapes(prs)

    print("Filling Slide 1...")
    fill_slide_1(prs)

    print("Filling Slide 2...")
    fill_slide_2(prs)

    print("Filling Slide 3...")
    fill_slide_3(prs)

    print("Filling Slide 4...")
    fill_slide_4(prs)

    print("Filling Slide 5...")
    fill_slide_5(prs)

    validate_final_deck(prs)

    prs.save(str(OUTPUT_PATH))

    print("\nPowerPoint generated successfully.")
    print(f"Output file: {OUTPUT_PATH}")


# ============================================================
# 6. RUN SCRIPT
# ============================================================

if __name__ == "__main__":
    try:
        build_deck()

    except Exception as error:
        print("\nPowerPoint generation failed.")
        print(f"Error type: {type(error).__name__}")
        print(f"Error details: {error}")
        raise