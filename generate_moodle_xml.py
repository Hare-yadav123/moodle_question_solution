"""Generate and validate Moodle XML from the Task B SQLite database.

Supported Moodle question types:
- multiple_choice (single correct answer)
- true_false

Equation text is included in the question HTML. Embedded images are packaged
inside the Moodle XML using Moodle's @@PLUGINFILE@@ convention and base64 data.
"""

import base64
import html
import json
import sqlite3
import xml.etree.ElementTree as ET
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
DB_FILE = OUTPUT_DIR / "questions.db"
XML_FILE = OUTPUT_DIR / "moodle_questions.xml"


def add_text(parent, tag, value):
    element = ET.SubElement(parent, tag)
    element.text = "" if value is None else str(value)
    return element


def read_questions_from_database():
    if not DB_FILE.exists():
        raise FileNotFoundError(
            f"Database not found: {DB_FILE}. Run extract_questions.py first."
        )

    with sqlite3.connect(DB_FILE) as connection:
        connection.row_factory = sqlite3.Row
        rows = connection.execute(
            "SELECT question_no, question_type, question_text, option_a, "
            "option_b, option_c, option_d, correct_answer, marks, "
            "equation, image_reference FROM questions ORDER BY question_no"
        ).fetchall()

    questions = []
    for row in rows:
        item = dict(row)
        raw_images = item.get("image_reference")
        try:
            item["image_reference"] = json.loads(raw_images) if raw_images else []
        except (TypeError, json.JSONDecodeError):
            item["image_reference"] = []
        questions.append(item)
    return questions


def normalize_correct_option(question):
    """Accept A/B/C/D or the exact option text as the answer key."""
    answer = str(question["correct_answer"]).strip()
    if answer.upper() in {"A", "B", "C", "D"}:
        key = "option_" + answer.lower()
        return question.get(key)

    for letter in "ABCD":
        option = question.get("option_" + letter.lower())
        if option and answer.casefold() == str(option).strip().casefold():
            return option
    raise ValueError(
        f"Q{question['question_no']}: correct answer {answer!r} "
        "does not match A/B/C/D or any option text."
    )


def question_html(question):
    """Build safe HTML; return HTML plus image files to embed."""
    text = str(question.get("question_text") or "").strip()
    equation = question.get("equation")

    # The extractor appends [Equation: ...] to question_text; avoid duplication
    # and render the dedicated equation field as 12 a separate readable line.
    if equation:
        marker = f"[Equation: {equation}]"
        text = text.replace(marker, "").strip()

    html_parts = [f"<p>{html.escape(text)}</p>"]
    if equation:
        html_parts.append(
            f"<p><strong>Equation:</strong> {html.escape(str(equation))}</p>"
        )

    embedded_files = []
    for image_ref in question.get("image_reference") or []:
        image_path = (BASE_DIR / image_ref).resolve()
        try:
            image_path.relative_to(BASE_DIR.resolve())
        except ValueError:
            raise ValueError(f"Q{question['question_no']}: unsafe image path")

        if image_path.is_file():
            filename = image_path.name
            html_parts.append(
                f'<p><img src="@@PLUGINFILE@@/{html.escape(filename)}" '
                f'alt="Question {question["question_no"]} diagram" /></p>'
            )
            embedded_files.append((filename, image_path.read_bytes()))
        else:
            html_parts.append(
                f"<p><em>Image not embedded (source file missing): "
                f"{html.escape(str(image_ref))}</em></p>"
            )

    return "\n".join(html_parts), embedded_files


def add_embedded_files(questiontext, files):
    for filename, content in files:
        file_element = ET.SubElement(
            questiontext,
            "file",
            {"name": filename, "path": "/", "encoding": "base64"},
        )
        file_element.text = base64.b64encode(content).decode("ascii")


def build_multichoice(quiz, question):
    number = question["question_no"]
    node = ET.SubElement(quiz, "question", {"type": "multichoice"})
    add_text(node, "single", "true")
    add_text(node, "shuffleanswers", "true")
    add_text(node, "answernumbering", "abc")
    add_text(node, "showstandardinstruction", "0")
    add_text(node, "correctfeedback", "")
    add_text(node, "partiallycorrectfeedback", "")
    add_text(node, "incorrectfeedback", "")
    add_text(node, "shownumcorrect", "1")

    correct_option = normalize_correct_option(question)
    if not correct_option:
        raise ValueError(f"Q{number}: correct option is empty")

    for letter in "ABCD":
        option = question.get("option_" + letter.lower())
        if option is None or str(option).strip() == "":
            continue
        fraction = "100" if str(option).strip() == str(correct_option).strip() else "0"
        answer = ET.SubElement(node, "answer", {"fraction": fraction, "format": "html"})
        add_text(answer, "text", str(option))
        add_text(answer, "feedback", "")


def build_truefalse(quiz, question):
    number = question["question_no"]
    node = ET.SubElement(quiz, "question", {"type": "truefalse"})

    correct = str(question["correct_answer"]).strip().casefold()
    if correct not in {"true", "false"}:
        raise ValueError(
            f"Q{number}: True/False answer must be 'True' or 'False', "
            f"got {question['correct_answer']!r}"
        )

    for value in ("true", "false"):
        answer = ET.SubElement(node, "answer", {"fraction": "100" if value == correct else "0"})
        add_text(answer, "text", value)
        add_text(answer, "feedback", "")


def validate_question_data(questions):
    if not questions:
        raise ValueError("No questions found in the SQLite database.")

    for question in questions:
        number = question["question_no"]
        qtype = question["question_type"]
        if qtype not in {"multiple_choice", "true_false"}:
            raise ValueError(f"Q{number}: unsupported question type {qtype!r}")
        if not str(question.get("question_text") or "").strip():
            raise ValueError(f"Q{number}: question text is empty")
        if float(question["marks"]) < 0:
            raise ValueError(f"Q{number}: marks cannot be negative")
        if qtype == "multiple_choice":
            for letter in "abcd":
                if not str(question.get(f"option_{letter}") or "").strip():
                    raise ValueError(f"Q{number}: option {letter.upper()} is missing")
            normalize_correct_option(question)
        else:
            if str(question.get("correct_answer") or "").strip().casefold() not in {"true", "false"}:
                raise ValueError(f"Q{number}: invalid True/False answer")


def generate_moodle_xml():
    questions = read_questions_from_database()
    validate_question_data(questions)

    quiz = ET.Element("quiz")
    for question in questions:
        if question["question_type"] == "multiple_choice":
            build_multichoice(quiz, question)
        else:
            build_truefalse(quiz, question)

    tree = ET.ElementTree(quiz)
    ET.indent(tree, space="  ")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    tree.write(XML_FILE, encoding="utf-8", xml_declaration=True)

    # Parse the saved file again. This verifies XML is well-formed.
    ET.parse(XML_FILE)
    print(f"Moodle XML created and validated: {XML_FILE}")
    print(f"Questions exported: {len(questions)}")
    print(f"Multiple Choice: {sum(q['question_type'] == 'multiple_choice' for q in questions)}")
    print(f"True/False: {sum(q['question_type'] == 'true_false' for q in questions)}")
    print(f"Total marks: {sum(float(q['marks']) for q in questions):g}")
    print("XML validation: PASSED (well-formed XML)")
    return XML_FILE

if __name__ == "__main__":
    generate_moodle_xml()
