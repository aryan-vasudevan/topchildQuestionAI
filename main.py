import openpyxl
from openai import OpenAI
import os
from dotenv import load_dotenv
from pydantic import BaseModel

class Question(BaseModel):
    question: str
    answerChoices: list[str]
    correctAnswer: str

class QuestionList(BaseModel):
    questions: list[Question]

def getNewQuestions(sampleQuestion):
    completion = client.beta.chat.completions.parse(
        model="gpt-4o-2024-08-06",
        messages=[{"role": "system", "content": PROMPT},
                  {"role": "user", "content": f"question - {sampleQuestion.question} \n answer choices - {sampleQuestion.answerChoices} \n correct answer - {sampleQuestion.correctAnswer}"}],
        response_format=Question
    )

    output = completion.choices[0].message.parsed
    print(output)
    print(type(output))

# Excel file
path = "questions.xlsx" # path to the excel file
wb = openpyxl.load_workbook(path) # load the workbook

sheet = wb.active

# API
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
PROMPT = os.getenv("PROMPT")
client = OpenAI(api_key=OPENAI_API_KEY)

questions = []
for row in sheet.iter_rows(min_row=2, min_col=2, max_row=20, max_col=8):
    question = str(row[0].value)
    answerChoices = [str(cell.value) for cell in row[1:6]]
    correctAnswer = str(row[6].value)

    questions.append(Question(question=question, answerChoices=answerChoices, correctAnswer=correctAnswer))

getNewQuestions(questions[0])

