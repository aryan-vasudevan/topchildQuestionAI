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
for cell in sheet['B']:
    if cell.value != None and cell != sheet['B1']:
        questions.append(cell.value)
