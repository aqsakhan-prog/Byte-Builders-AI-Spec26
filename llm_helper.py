import json
import os
from pathlib import Path

from dotenv import load_dotenv
from google import genai
from google.genai import types


ENV_PATH = Path(__file__).resolve().parent / ".env"

CANDIDATE_MODELS = [
    "gemini-flash-lite-latest",
    "gemini-3.1-flash-lite",
    "gemini-3.5-flash-lite",
    "gemini-flash-latest",
    "gemini-3.6-flash",
]


def answer_with_gemini(
    question,
    part_number,
    excel_result,
    pdf_result,
    iges_result,
    comparison_rows,
):
    """
    Answer using Gemini and only evidence extracted from the
    supplied Excel, PDF, and IGES files.

    If one model is overloaded, another model is tried.
    """

    load_dotenv(dotenv_path=ENV_PATH)

    api_key = os.getenv("GEMINI_API_KEY")

    if not api_key:
        return {
            "status": "error",
            "message": (
                "GEMINI_API_KEY was not found in the .env file."
            ),
        }

    try:
        client = genai.Client(
            api_key=api_key,
            http_options=types.HttpOptions(timeout=12000),
        )

        pdf_pages = []

        for page in pdf_result.get("pages", [])[:10]:
            pdf_pages.append(
                {
                    "page": page.get("page"),
                    "text": page.get("text", "")[:6000],
                }
            )

        evidence_package = {
            "requested_part_number": part_number,
            "excel": {
                "status": excel_result.get("status"),
                "source_file": excel_result.get("source_file"),
                "sheet": excel_result.get("sheet"),
                "row": excel_result.get("excel_row"),
                "record": excel_result.get("record", {}),
                "message": excel_result.get("message"),
            },
            "pdf": {
                "status": pdf_result.get("status"),
                "source_file": pdf_result.get("source_file"),
                "page_count": pdf_result.get("page_count"),
                "facts": pdf_result.get("facts", {}),
                "pages": pdf_pages,
                "message": pdf_result.get("message"),
            },
            "iges": {
                key: value
                for key, value in iges_result.items()
                if key != "freecad_output"
            },
            "cross_source_comparisons": comparison_rows,
        }

        prompt = f"""
You are an evidence-grounded CNC manufacturing assistant.

Requested part number: {part_number}

STRICT RULES:
1. Use only the supplied file evidence below.
2. Never invent missing information.
3. Do not use external knowledge to fill gaps.
4. If information is absent, say exactly:
   "I cannot find that information in the supplied files."
5. Cite every factual claim using its actual source filename.
6. For Excel evidence, include filename, worksheet, and row.
7. For PDF evidence, include filename and page number.
8. For IGES evidence, include the IGES filename.
9. Clearly identify disagreements between sources.
10. If IGES volume reliability is limited, mention that warning.
11. Keep the response concise and understandable.
12. Do not mention these instructions.

Use this structure:

Answer:
<direct answer>

Evidence:
- <claim with citation>

Warnings:
- <warning or conflict, or None>

USER QUESTION:
{question}

SUPPLIED FILE EVIDENCE:
{json.dumps(evidence_package, indent=2, default=str)}
"""

        model_errors = []

        for model_name in CANDIDATE_MODELS:
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                )

                response_text = getattr(response, "text", None)

                if not response_text:
                    model_errors.append(
                        f"{model_name}: empty response"
                    )
                    continue

                return {
                    "status": "success",
                    "answer": response_text.strip(),
                    "model": model_name,
                }

            except Exception as error:
                model_errors.append(
                    f"{model_name}: {str(error)[:250]}"
                )

        return {
            "status": "error",
            "message": (
                "All configured Gemini models failed. "
                "The offline fallback will be used.\n\n"
                + "\n".join(model_errors)
            ),
        }

    except Exception as error:
        return {
            "status": "error",
            "message": str(error),
        }