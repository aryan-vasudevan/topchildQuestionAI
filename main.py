import openpyxl
import os
import json
from openai import OpenAI
from dotenv import load_dotenv
from pydantic import BaseModel

# Initialize API
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
INSTRUCTIONS = os.getenv("INSTRUCTIONS")

client = OpenAI(api_key=OPENAI_API_KEY)

# Question model for structured outputs
class Question(BaseModel):
    questionText: str
    answerChoices: list[str]
    correctAnswer: str

# Get new question
thread = client.beta.threads.create()
def getNewQuestion(sampleQuestion, questionJSONList):

    message = client.beta.threads.messages.create(
        thread_id=thread.id,
        role="user",
        content=f"question - {sampleQuestion.questionText} \n answer choices - {sampleQuestion.answerChoices} \n correct answer - {sampleQuestion.correctAnswer}"
    )

    run = client.beta.threads.runs.create_and_poll(
        thread_id=thread.id,
        assistant_id="asst_JvXYRQsCuulfT0rGke7VZA0D",
    )

    if run.status == "completed": 
        messages = list(client.beta.threads.messages.list(
            thread_id=thread.id,
    ))
    questionJSON = messages[0].content[0].text.value
    print(questionJSON)
    questionJSONList.append(questionJSON)
# Excel file
path = "questions.xlsx"
wb = openpyxl.load_workbook(path)

# Sample question sheet
sheet = wb.active

# New question sheet
wb.create_sheet("Sheet 2")
sheet2 = wb["Sheet 2"]

# Read questions
sampleQuestionList = []
for row in sheet.iter_rows(min_row=2, min_col=2, max_row=20, max_col=8):
    # Read specific cells
    questionText = str(row[0].value)
    answerChoices = [str(cell.value) for cell in row[1:6]]
    correctAnswer = str(row[6].value)

    # Keep questions in model form to keep it organized
    sampleQuestionList.append(Question(questionText=questionText, answerChoices=answerChoices, correctAnswer=correctAnswer))

# Get a new question for each sample question
questionJSONList = []
for sampleQuestion in sampleQuestionList:
    getNewQuestion(sampleQuestion, questionJSONList)

newQuestionList = []
for questionJSON in questionJSONList:
    questionObj = json.loads(questionJSON)

    newQuestion = Question(questionText=questionObj["questionText"], answerChoices=questionObj["answerChoices"], correctAnswer=questionObj["correctAnswer"])
    newQuestionList.append(newQuestion)
