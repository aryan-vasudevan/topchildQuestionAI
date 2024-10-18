import openpyxl
from openai import OpenAI
import os
from dotenv import load_dotenv
from pydantic import BaseModel

class Question(BaseModel):
    question: str
    answerChoices: str
    correctAnswer: str

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

response = client.chat.completions.create(
    model="gpt-4o-2024-08-06",
    messages=[
        {
            "role": "system", 
            "content": "You extract email addresses into JSON data."
        },
        {
            "role": "user", 
            "content": "Feeling stuck? Send a message to help@mycompany.com."
        }
    ],
    response_format={
        "type": "json_schema",
        "json_schema": {
            "name": "email_schema",
            "schema": {
                "type": "object",
                "properties": {
                    "email": {
                        "description": "The email address that appears in the input",
                        "type": "string"
                    },
                    "additionalProperties": False
                }
            }
        }
    }
)

print(response.choices[0].message.content);
