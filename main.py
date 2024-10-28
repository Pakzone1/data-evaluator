import os
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
from openai import OpenAI
import pandas as pd
import json
from typing import List

app = FastAPI()

client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

ASSISTANT_ID = "asst_K32nPAuLJ8Vkdo5x9PKBxPR5"

def excel_to_json(file_path: str, sheet_name: str = "Table 1") -> str:
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    df[['Chapter', 'Section', 'Sub-Section']] = df[['Chapter', 'Section', 'Sub-Section']].ffill()

    json_data = df.to_json(orient="records", indent=4)

    return json_data

@app.post("/process_excel/")
async def process_excel(file: UploadFile = File(...)):

    temp_file_path = f"temp_{file.filename}"
    with open(temp_file_path, "wb") as buffer:
        buffer.write(await file.read())

    try:

        json_data = excel_to_json(temp_file_path)

        thread = client.beta.threads.create()

        message = client.beta.threads.messages.create(
            thread_id=thread.id,
            role="user",
            content=f"{json_data}"
        )

        run = client.beta.threads.runs.create(
            thread_id=thread.id,
            assistant_id=ASSISTANT_ID
        )

        while run.status != "completed":
            run = client.beta.threads.runs.retrieve(
                thread_id=thread.id,
                run_id=run.id
            )

        messages = client.beta.threads.messages.list(thread_id=thread.id)
        assistant_response = next(msg.content[0].text.value for msg in messages if msg.role == "assistant")

        print("Assistant's response:")
        print(assistant_response)

        return JSONResponse(content={assistant_response})

    except Exception as e:
        return JSONResponse(content={"error": str(e)}, status_code=500)

    finally:

        os.remove(temp_file_path)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
