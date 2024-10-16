import openpyxl
from openai import OpenAI
import os
from dotenv import load_dotenv

class Question():
    def __init__(self, question, answerChoices, correctAnswer):
        self.question = question
        self.answerChoices = answerChoices
        self.correctAnswer = correctAnswer

# Excel file
path = "questions.xlsx" # path to the excel file
wb = openpyxl.load_workbook(path) # load the workbook

sheet = wb.active  

# API
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=OPENAI_API_KEY)

questions = []
for row in sheet.iter_rows(min_row=2, min_col=2, max_row=20, max_col=8):
    question = row[0].value
    answerChoices = [cell.value for cell in row[1:6]]
    correctAnswer = row[6].value
    questions.append(Question(question, answerChoices, correctAnswer))

print(questions[0].question, questions[0].answerChoices, questions[0].correctAnswer)
