"""
Extract questions, options, answers, equations and embedded images from a DOCX file.
Outputs:
    output/questions.csv
    output/questions.json
    output/images/*
Install:
    pip install python-docx
Expected input:
    input/Sample_Question_Bank_Source.docx
"""
import csv
import hashlib
import json
import re
from pathlib import Path
import sqlite3
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table


# ---------------------------------------------------------------------
# 1. Paths and settings
# ---------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
INPUT_FILE = BASE_DIR / "input" / "Sample_Question_Bank_Source.docx"
OUTPUT_DIR = BASE_DIR / "output"
IMAGE_DIR = OUTPUT_DIR / "images"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
IMAGE_DIR.mkdir(parents=True, exist_ok=True)

EXPECTED_QUESTIONS = 10

COLUMNS = [
    "Question No.", "Question Type", "Question Text",
    "Option A", "Option B", "Option C", "Option D",
    "Correct Answer", "Marks", "Difficulty",
]

HEADER_RE = re.compile(
    r"Q\s*(\d+)\s*\.\s*"
    r"\(\s*(Multiple\s+Choice|True\s*/\s*False)\s*"
    r"\|\s*Marks\s*:\s*(\d+)\s*"
    r"\|\s*Difficulty\s*:\s*(Easy|Medium|Hard)\s*\)",
    re.IGNORECASE,
)
OPTION_RE = re.compile(r"^\s*([A-D])\s*[.)]\s*(.*?)\s*$", re.IGNORECASE)
ANSWER_RE = re.compile(r"\bCorrect\s+Answer\s*:\s*(.*)", re.IGNORECASE)
EQUATION_RE = re.compile(r"\bEquation\s*:\s*(.*)", re.IGNORECASE)


# ---------------------------------------------------------------------
# 2. Read document blocks in their original order (paragraphs + tables)
# ---------------------------------------------------------------------

def iter_block_items(parent):
    """Yield Paragraph and Table objects in document order."""
    if hasattr(parent, "_tc"):       # table cell
        parent_element = parent._tc
    elif hasattr(parent, "element"): # Document
        parent_element = parent.element.body
    else:
        parent_element = parent

    for child in parent_element.iterchildren():
        # print("*"*12)
        # print(child.tag)
        # print("*"*12)
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
            # print("paragraph",Paragraph(child, parent))
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def extract_images(paragraph):
    """Save images embedded in a paragraph and return relative paths."""
    saved_paths = []

    for blip in paragraph._p.xpath(".//a:blip"):
        relationship_id = blip.get(qn("r:embed"))
        if not relationship_id:
            continue

        image_part = paragraph.part.related_parts.get(relationship_id)
        if image_part is None:
            continue

        image = image_part.image
        digest = hashlib.sha256(image.blob).hexdigest()[:12]
        filename = f"image_{digest}.{image.ext}"
        image_path = IMAGE_DIR / filename

        if not image_path.exists():
            image_path.write_bytes(image.blob)

        relative_path = image_path.relative_to(BASE_DIR).as_posix()
        if relative_path not in saved_paths:
            saved_paths.append(relative_path)

    return saved_paths


def read_document(parent):
    """Return ordered records; table cell paragraphs are also supported."""
    records = []
    for block in iter_block_items(parent):
        if isinstance(block, Paragraph):
            records.append({
                "text": re.sub(r"\s+", " ", block.text).strip(),
                "images": extract_images(block),
            })
            
        else:
            for row in block.rows:
                for cell in row.cells:
                    records.extend(read_document(cell))

    # print("*"*20)
    # print("records all: ",records)
    # print("*"*20)
    return records

# ---------------------------------------------------------------------
# 3. Question data helpers
# ---------------------------------------------------------------------

def create_question(match):
    raw_type = re.sub(r"\s+", " ", match.group(2)).lower()
    question_type = (
        "multiple_choice" if raw_type == "multiple choice" else "true_false"
    )

    question = {
        "question_no": int(match.group(1)),
        "question_type": question_type,
        "question_text": "",
        "option_a": None,
        "option_b": None,
        "option_c": None,
        "option_d": None,
        "correct_answer": "",
        "marks": int(match.group(3)),
        "difficulty": match.group(4).capitalize(),
        "equation": None,
        "image_reference": [],
    }

    if question_type == "true_false":
        question["option_a"] = "True"
        question["option_b"] = "False"

    # print("question : ",question)
    return question


def append_question_text(question, text):
    text = text.strip()
    if not text:
        return

    print("debug question",question)
    print("debug question type,",type(question))
    print("debug text",text)
    if question is None:
        raise ValueError(
            "question object is none , inside append_question_text function"
        )
    if question["question_text"]:
        question["question_text"] += " " + text
    else:
        question["question_text"] = text


def process_question_content(question, text):
    """
    Process a paragraph's text. A paragraph may contain question text,
    an option, Equation:, or Correct Answer: on the same line.
    """
    print("\n--- PROCESS QUESTION DEBUG ---")
    print("Question:", question)
    print("Question Type:", type(question))
    print("Text:", text)

    text = text.strip()
    # print("text : ",text)
    if not text:
        return

    # Correct answer may appear after other text in the same paragraph.
    answer_match = ANSWER_RE.search(text)
    if answer_match:
        before_answer = text[:answer_match.start()].strip()
        if before_answer:
            process_question_content(question, before_answer)
        question["correct_answer"] = answer_match.group(1).strip()
        return

    equation_match = EQUATION_RE.search(text)
    if equation_match:
        before_equation = text[:equation_match.start()].strip()
        if before_equation:
            process_question_content(question, before_equation)
        question["equation"] = equation_match.group(1).strip()
        return

    option_match = OPTION_RE.match(text)
    if option_match:
        key = "option_" + option_match.group(1).lower()
        question[key] = option_match.group(2).strip()
        return

    if text.lower().startswith(("instructions:", "end of source")):
        return

    # print("question text : ",question["question_text"]) 
    append_question_text(question, text)


# ---------------------------------------------------------------------
# 4. Extract questions
# ---------------------------------------------------------------------

def extract_questions(records):
    questions = []
    current_question = None

    for record in records:
        line = record["text"]

        # Search the actual DOCX paragraph, not a hard-coded test string.
        header_match = HEADER_RE.search(line)

        if header_match:
            if current_question is not None:
                questions.append(current_question)

            current_question = create_question(header_match)

            # Important: some DOCX files put Q number, question text and
            # answer in one paragraph (for example Q2 and Q4 in this file).
            remaining_text = line[header_match.end():].strip()
            if remaining_text:
                print("Current Question:", current_question)
                print("Current Question Type:", type(current_question))
                print("Remaining Text:", remaining_text)
                process_question_content(current_question, remaining_text)

        elif current_question is not None:
            process_question_content(current_question, line)

        # Attach images even when the paragraph contains no text.
        if current_question is not None:
            for image_path in record["images"]:
                if image_path not in current_question["image_reference"]:
                    current_question["image_reference"].append(image_path)

    if current_question is not None:
        questions.append(current_question)

    # print("q1",questions)
    return questions


# ---------------------------------------------------------------------
# 5. Validate extracted data
# ---------------------------------------------------------------------

def validate_questions(questions):
    if len(questions) != EXPECTED_QUESTIONS:
        found = [q["question_no"] for q in questions]
        raise ValueError(
            f"Expected {EXPECTED_QUESTIONS} questions, found {len(questions)}. "
            f"Detected question numbers: {found}. "
            "Check that question headings use the format "
            "'Q1. (Multiple Choice | Marks: 1 | Difficulty: Easy)'."
        )

    seen = set()

    for question in questions:
        number = question["question_no"]

        if number in seen:
            raise ValueError(f"Duplicate question number: Q{number}")
        seen.add(number)

        if not question["question_text"].strip():
            raise ValueError(f"Question text missing: Q{number}")

        if not question["correct_answer"].strip():
            raise ValueError(f"Correct answer missing: Q{number}")

        if question["question_type"] == "multiple_choice":
            for key in ("option_a", "option_b", "option_c", "option_d"):
                if not question[key]:
                    raise ValueError(f"{key} missing in Q{number}")

        # Keep equation and image information in the exported question text
        # as well as in their dedicated JSON fields.
        if question["equation"]:
            marker = f"[Equation: {question['equation']}]"
            if marker not in question["question_text"]:
                question["question_text"] += " " + marker

        if question["image_reference"]:
            marker = "[Diagram: " + ", ".join(question["image_reference"]) + "]"
            if marker not in question["question_text"]:
                question["question_text"] += " " + marker

    questions.sort(key=lambda item: item["question_no"])

    expected_numbers = list(range(1, EXPECTED_QUESTIONS + 1))
    actual_numbers = [q["question_no"] for q in questions]
    if actual_numbers != expected_numbers:
        raise ValueError(
            f"Question numbering mismatch. Expected {expected_numbers}, "
            f"found {actual_numbers}"
        )


# ---------------------------------------------------------------------
# 6. Save CSV and JSON
# ---------------------------------------------------------------------

def save_csv(questions):
    csv_path = OUTPUT_DIR / "questions.csv"

    with csv_path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=COLUMNS)
        writer.writeheader()

        for q in questions:
            writer.writerow({
                "Question No.": q["question_no"],
                "Question Type": q["question_type"],
                "Question Text": q["question_text"],
                "Option A": q["option_a"] or "",
                "Option B": q["option_b"] or "",
                "Option C": q["option_c"] or "",
                "Option D": q["option_d"] or "",
                "Correct Answer": q["correct_answer"],
                "Marks": q["marks"],
                "Difficulty": q["difficulty"],
            })

    print(f"CSV file created: {csv_path}")


def save_json(questions):
    json_path = OUTPUT_DIR / "questions.json"

    with json_path.open("w", encoding="utf-8") as file:
        json.dump(questions, file, indent=4, ensure_ascii=False)

    print(f"JSON file created: {json_path}")


# ---------------------------------------------------------------------
# 7. SQLite database: Task B
# ---------------------------------------------------------------------

DATABASE_FILE = OUTPUT_DIR / "questions.db"


def create_database():
    """Create SQLite database and its questions table if needed."""
    with sqlite3.connect(DATABASE_FILE) as connection:
        connection.execute("""
            CREATE TABLE IF NOT EXISTS questions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                question_no INTEGER NOT NULL UNIQUE,
                question_type TEXT NOT NULL,
                question_text TEXT NOT NULL,
                option_a TEXT,
                option_b TEXT,
                option_c TEXT,
                option_d TEXT,
                correct_answer TEXT NOT NULL,
                marks INTEGER NOT NULL CHECK (marks >= 0),
                difficulty TEXT NOT NULL
                    CHECK (difficulty IN ('Easy', 'Medium', 'Hard')),
                equation TEXT,
                image_reference TEXT,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """)
        connection.commit()


def save_to_database(questions):
    """Save new questions and update existing question numbers."""
    create_database()
    rows = [
        (
            q["question_no"], q["question_type"], q["question_text"],
            q["option_a"], q["option_b"], q["option_c"], q["option_d"],
            q["correct_answer"], q["marks"], q["difficulty"], q["equation"],
            json.dumps(q["image_reference"], ensure_ascii=False),
        )
        for q in questions
    ]

    with sqlite3.connect(DATABASE_FILE) as connection:
        connection.executemany("""
            INSERT INTO questions (
                question_no, question_type, question_text,
                option_a, option_b, option_c, option_d,
                correct_answer, marks, difficulty, equation, image_reference
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(question_no) DO UPDATE SET
                question_type = excluded.question_type,
                question_text = excluded.question_text,
                option_a = excluded.option_a,
                option_b = excluded.option_b,
                option_c = excluded.option_c,
                option_d = excluded.option_d,
                correct_answer = excluded.correct_answer,
                marks = excluded.marks,
                difficulty = excluded.difficulty,
                equation = excluded.equation,
                image_reference = excluded.image_reference
        """, rows)
        connection.commit()

    print(f"SQLite database updated: {DATABASE_FILE}")
    print(f"Questions saved/updated: {len(rows)}")

# ---------------------------------------------------------------------
# 7. Main
# ---------------------------------------------------------------------

def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Word file not found: {INPUT_FILE}\n"
            "Place the DOCX file inside the project's 'input' folder."
        )

    print("Reading Word document...")
    document = Document(INPUT_FILE)

    print("Reading paragraphs and tables...")
    records = read_document(document)
    print(f"Text/image records found: {len(records)}")

    print("Extracting questions and images...")
    questions = extract_questions(records)
    print(f"Questions detected: {len(questions)}")

    validate_questions(questions)

    save_csv(questions)
    save_json(questions)
    save_to_database(questions)
    print("docx question updated in database.")

    image_count = sum(len(q["image_reference"]) for q in questions)

    print("\nExtraction completed successfully!")
    print("Total questions:", len(questions))
    print("Images extracted:", image_count)
    print("Output folder:", OUTPUT_DIR)


if __name__ == "__main__":
    main()
