-- 1. List all questions in question-number order
-- select question_no, question_type, question_text, marks, difficulty
select * from questions order by question_no asc;

-- 2. Count questions by difficulty
select difficulty, COUNT(difficulty) as question_count from questions
group by difficulty order by case difficulty when 'Easy' then 1 when 'Medium' then 2
when 'Hard' then 3 else 4 end;

-- 3. Calculate total available marks
select SUM(marks) as total_available_marks
from questions;

-- 4. List questions worth 2 or more marks
select question_no, question_text, marks, difficulty
from questions where marks >= 2 order by  question_no asc;

