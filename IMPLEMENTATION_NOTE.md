Implementation Note — Moodle Question Bank Automation

I normalized the ten supplied sample questions into a common record structure and generated CSV, SQLite and 
Moodle XML outputs. The workflow serptates source extraction, normalized question data, database sorrages and XML generation so each stage can be tested independently. 

Libraries: Python standard library modules csv, json, sqlite3, pathlib and xml.etree.ElementTree are used for data handling, SQL storage and XML generation. python-docx is used by the optional DOCX extractor.

Assumptions: The assignment specifies a Word DOCX document, and the provided source document is also in DOCX format. Therefore, the Python script extracts the question content directly from the DOCX file, including paragraphs, tables, and embedded images. The extracted information is organized into structured Python data and then stored in an SQLite database.
For True/False questions, Option A represents True and Option B represents False, while Options C and D remain empty.
  
Equation and diagram: Q9's equation is retained as readable text (2x + 6 = 14) and appended to the Moodle question text. Q10's diagram is represented by its visible sequence: Input (Student data) -> Python (Process) -> Output (XML file). The original diagram image is not embedded in the XML.

Validation and improvements: the generated XML is parsed with ElementTree to confirm well-formedness. With more time, I would add stronger parsing for varied Word layouts, preserve embedded images as Moodle files, add automated tests, and import the XML into the target Moodle version to verify version-specific compatibility.
