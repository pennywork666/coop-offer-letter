from __future__ import annotations

import base64
from copy import deepcopy
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
import re

import streamlit as st
from docx import Document


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "Offer Letter Template.docx"
LOGO_PATH = BASE_DIR / "Midea.png"
LOCATION_DETAILS = {
    "Louisville": "2700 Chestnut Station Ct, Louisville, KY 40299",
    "Boston": "260 Charles Street, Suite 401, Waltham, MA 02453",
}


@dataclass
class OfferLetterData:
    candidate_name: str
    letter_date: date
    position_title: str
    work_location: str
    employment_start_date: date
    employment_end_date: date
    hourly_rate: Decimal
    relocation_assistance: Decimal
    output_stem: str


def format_long_date(value: date) -> str:
    return f"{value.strftime('%B')} {value.day}, {value.year}"


def format_money(value: Decimal) -> str:
    normalized = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{normalized:,.2f}"


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\\\|?*]+', "", value).strip()
    cleaned = re.sub(r"\s+", "_", cleaned)
    cleaned = re.sub(r"_+", "_", cleaned)
    return cleaned.strip("._") or "offer_letter"


def compute_overtime_rate(hourly_rate: Decimal) -> Decimal:
    return hourly_rate * Decimal("1.5")


def get_image_data_uri(image_path: Path) -> str:
    encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def build_default_output_stem(candidate_name: str) -> str:
    return sanitize_filename(f"{candidate_name} Offer Letter")


def extract_reporting_clause(paragraph_text: str) -> str:
    match = re.search(
        r"You will be reporting to\s*(.*?)\.\s*For purposes",
        paragraph_text,
        flags=re.DOTALL,
    )
    if not match:
        return ""

    clause = " ".join(match.group(1).split()).strip(" ,.")
    return clause


def clear_runs(paragraph) -> None:
    for child in list(paragraph._p):
        if child.tag.endswith("}r") or child.tag.endswith("}hyperlink"):
            paragraph._p.remove(child)


def remove_paragraph(paragraph) -> None:
    paragraph._element.getparent().remove(paragraph._element)


def add_styled_run(paragraph, text: str, source_run=None) -> None:
    run = paragraph.add_run(text)
    if source_run is not None and source_run._r.rPr is not None:
        run._r.insert(0, deepcopy(source_run._r.rPr))


def replace_paragraph(paragraph, parts: list[tuple[str, int | None]]) -> None:
    original_runs = list(paragraph.runs)
    clear_runs(paragraph)

    if not parts:
        paragraph.add_run("")
        return

    for text, run_index in parts:
        source_run = None
        if original_runs:
            if run_index is not None and 0 <= run_index < len(original_runs):
                source_run = original_runs[run_index]
            else:
                source_run = original_runs[0]
        add_styled_run(paragraph, text, source_run)


def build_offer_letter_document(template_path: Path, data: OfferLetterData) -> Document:
    document = Document(template_path)
    paragraphs = document.paragraphs
    location_text = LOCATION_DETAILS.get(data.work_location, data.work_location)
    overtime_rate = compute_overtime_rate(data.hourly_rate)
    relocation_paragraph = paragraphs[15]
    reporting_clause = extract_reporting_clause(paragraphs[8].text)
    reporting_sentence = f" You will be reporting to {reporting_clause}." if reporting_clause else ""

    if len(paragraphs) < 57:
        raise ValueError("The template structure changed. Expected at least 57 paragraphs.")

    replace_paragraph(paragraphs[6], [(format_long_date(data.letter_date), None)])
    replace_paragraph(
        paragraphs[7],
        [
            ("Dear ", 0),
            (data.candidate_name, 2),
            (",", 3),
        ],
    )
    replace_paragraph(
        paragraphs[8],
        [
            (
                (
                    f'On behalf of Midea America Corp. ("Midea"), I am pleased to offer you '
                    f'the full-time position {data.position_title} Co-Op, working at '
                    f'{location_text}.{reporting_sentence} For purposes of this letter, '
                    f'your first day of work at Midea will be considered your "Employment Start '
                    f'Date." Your Employment Start Date will be on '
                    f'{format_long_date(data.employment_start_date)} until '
                    f'{format_long_date(data.employment_end_date)}.'
                ),
                0,
            )
        ],
    )
    replace_paragraph(
        paragraphs[14],
        [
            (
                (
                    "Base Salary (non-exempt): Your hourly rate will be "
                    f"${format_money(data.hourly_rate)} per hour, and in the event you are "
                    "authorized to work overtime, you will be paid out at "
                    f"${format_money(overtime_rate)} per hour; paid semi-monthly, subject "
                    "to annual review. Less applicable taxes, deductions and withholdings. "
                    "Average work week will be 40 hours, as approved, overtime may be needed."
                ),
                0,
            )
        ],
    )
    if data.relocation_assistance != Decimal("0"):
        replace_paragraph(
            relocation_paragraph,
            [
                (
                    (
                        "Relocation Assistance: In order to assist in our transition and move "
                        "from your current location to "
                        f"{location_text} we will provide relocation assistance of "
                        f"${format_money(data.relocation_assistance)} on a grossed-up, pre-tax "
                        'basis (the "Relocation Benefit") which is payable during the first pay '
                        "period, to be paid within thirty (30) days of employment. The "
                        "Relocation Benefit will be subject to full repayment to Midea if you "
                        "voluntarily resign from Midea or are terminated for cause prior to the "
                        "12 months of your Employment Start Date."
                    ),
                    0,
                )
            ],
        )
    replace_paragraph(
        paragraphs[52],
        [
            ("Name (Please Print): ", 0),
            (data.candidate_name, 0),
        ],
    )
    replace_paragraph(
        paragraphs[56],
        [
            ("Planned Employment Start Date: ", 0),
            (format_long_date(data.employment_start_date), 0),
        ],
    )

    if data.relocation_assistance == Decimal("0"):
        remove_paragraph(relocation_paragraph)

    return document


def build_offer_letter_bytes(template_path: Path, data: OfferLetterData) -> bytes:
    document = build_offer_letter_document(template_path, data)
    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output.getvalue()
def build_data(
    candidate_name: str,
    letter_date: date,
    position_title: str,
    work_location: str,
    employment_start_date: date,
    employment_end_date: date,
    hourly_rate: float | int,
    relocation_assistance: float | int,
    output_stem: str,
) -> OfferLetterData:
    return OfferLetterData(
        candidate_name=candidate_name.strip(),
        letter_date=letter_date,
        position_title=position_title.strip(),
        work_location=work_location.strip(),
        employment_start_date=employment_start_date,
        employment_end_date=employment_end_date,
        hourly_rate=Decimal(str(hourly_rate)),
        relocation_assistance=Decimal(str(relocation_assistance)),
        output_stem=sanitize_filename(output_stem),
    )


def main() -> None:
    st.set_page_config(page_title="MARC CO-OP Offer Letter Generator", layout="wide")
    logo_markup = ""
    if LOGO_PATH.exists():
        logo_markup = f'<img class="midea-logo" src="{get_image_data_uri(LOGO_PATH)}" alt="Midea logo" />'

    st.markdown(
        f"""
        <style>
        [data-testid="stAppViewContainer"] {{
            background: linear-gradient(180deg, #dff1ff 0%, #c7e4ff 100%);
        }}
        [data-testid="stHeader"] {{
            background: transparent;
        }}
        .main .block-container {{
            max-width: 1180px;
            padding-top: 3.5rem;
            padding-bottom: 2.5rem;
        }}
        div[data-testid="stForm"] {{
            background: rgba(255, 255, 255, 0.98);
            border: 1px solid rgba(104, 157, 214, 0.22);
            border-radius: 18px;
            padding: 1.25rem;
            box-shadow: 0 16px 36px rgba(72, 120, 174, 0.12);
        }}
        .midea-hero {{
            position: relative;
            min-height: 88px;
            margin-bottom: 1.75rem;
        }}
        .midea-logo {{
            position: absolute;
            top: 0;
            left: 0;
            width: 180px;
            height: auto;
        }}
        h1.marc-title {{
            text-align: center;
            color: #0f4d8a;
            margin: 0.75rem 0 2rem;
            letter-spacing: 0.02em;
        }}
        div[data-testid="stDownloadButton"] {{
            display: flex;
            justify-content: center;
            width: 100%;
        }}
        div[data-testid="stDownloadButton"] button {{
            width: 340px;
            min-height: 4rem;
            font-size: 1.18rem;
            font-weight: 600;
            border-radius: 12px;
        }}
        </style>
        <div class="midea-hero">
            {logo_markup}
            <h1 class="marc-title">MARC CO-OP Offer Letter Generator</h1>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not TEMPLATE_PATH.exists():
        st.error(f"Template not found: {TEMPLATE_PATH}")
        return

    today = date.today()

    first_left, first_right = st.columns(2)
    with first_left:
        candidate_name = st.text_input("Name", value="")
    with first_right:
        position_title = st.text_input("Job title (do not include 'Co-Op')", value="")

    second_left, second_right = st.columns(2)
    with second_left:
        employment_start_date = st.date_input("Employment start date", value=None)
    with second_right:
        employment_end_date = st.date_input("Employment end date", value=None)

    third_left, third_right = st.columns(2)
    with third_left:
        hourly_rate = st.number_input(
            "Hourly rate ($)",
            min_value=0.0,
            value=None,
            step=0.5,
            placeholder="Enter hourly rate",
        )
    with third_right:
        relocation_assistance = st.number_input(
            "Relocation assistance ($)",
            min_value=0.0,
            value=None,
            step=100.0,
            placeholder="Enter relocation assistance",
        )

    fourth_left, fourth_right = st.columns(2)
    with fourth_left:
        work_location = st.selectbox(
            "Working location",
            options=list(LOCATION_DETAILS.keys()),
            index=None,
            placeholder="Select a location",
        )
    with fourth_right:
        st.empty()

    required_fields = {
        "Name": candidate_name,
        "Job title": position_title,
        "Working location": work_location or "",
    }

    if relocation_assistance is None:
        relocation_assistance = 0.0

    validation_errors = [label for label, value in required_fields.items() if not value.strip()]
    if employment_start_date is None or employment_end_date is None:
        validation_errors.append("Employment dates")
    if hourly_rate is None:
        validation_errors.append("Hourly rate")
    if (
        employment_start_date is not None
        and employment_end_date is not None
        and employment_end_date < employment_start_date
    ):
        validation_errors.append("Employment end date must be on or after the employment start date")

    final_output_stem = build_default_output_stem(candidate_name)
    download_data = b""
    generation_error = None

    if not validation_errors:
        try:
            offer_data = build_data(
                candidate_name=candidate_name,
                letter_date=today,
                position_title=position_title,
                work_location=work_location,
                employment_start_date=employment_start_date,
                employment_end_date=employment_end_date,
                hourly_rate=hourly_rate,
                relocation_assistance=relocation_assistance,
                output_stem=final_output_stem,
            )
            download_data = build_offer_letter_bytes(TEMPLATE_PATH, offer_data)
        except Exception as exc:
            generation_error = str(exc)

    if generation_error:
        st.error(f"Generation failed: {generation_error}")
    elif validation_errors:
        message = ", ".join(validation_errors)
        st.caption(f"Fill in the required fields before generating: {message}")

    st.download_button(
        label="Generate offer letter",
        data=download_data,
        file_name=f"{final_output_stem}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        disabled=bool(validation_errors or generation_error),
    )


if __name__ == "__main__":
    main()
