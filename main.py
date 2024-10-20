import openpyxl
import os
from openai import OpenAI
from dotenv import load_dotenv
from pydantic import BaseModel

# API
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PROMPT = os.getenv("PROMPT")

client = OpenAI(api_key=OPENAI_API_KEY)

# Question model for structured outputs
class Question(BaseModel):
    questionText: str
    answerChoices: list[str]
    correctAnswer: str

# Get new question
def getNewQuestion(sampleQuestion):
    completion = client.beta.chat.completions.parse(
        model="gpt-4o-2024-08-06",
        messages=[{"role": "system", "content": PROMPT},
                  {"role": "user", "content": f"question - {sampleQuestion.questionText} \n answer choices - {sampleQuestion.answerChoices} \n correct answer - {sampleQuestion.correctAnswer}"}],
        response_format=Question
    )

    output = completion.choices[0].message.parsed
    return output

# Excel file
path = "questions.xlsx" # path to the excel file
wb = openpyxl.load_workbook(path) # load the workbook

sheet = wb.active

# Read questions
sampleQuestions = []
for row in sheet.iter_rows(min_row=2, min_col=2, max_row=20, max_col=8):
    questionText = str(row[0].value)
    answerChoices = [str(cell.value) for cell in row[1:6]]
    correctAnswer = str(row[6].value)

    sampleQuestions.append(Question(questionText=questionText, answerChoices=answerChoices, correctAnswer=correctAnswer))

for sq in sampleQuestions:
    newQuestion = getNewQuestion(sq)
    print(f"\nquestion - {newQuestion.questionText}\nanswerChoices - {newQuestion.answerChoices} \ncorrectAnswer - {newQuestion.correctAnswer}")

# Create new sheet
wb.create_sheet("sheet 2")