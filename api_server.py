# api_server.py
from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import RedirectResponse
from fastapi.middleware.cors import CORSMiddleware
import uuid
import json
import time
from typing import Dict, Any
import os

app = FastAPI(title="Letter Formatter API")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Simple in-memory storage (in production, use Redis or a database)
letter_storage: Dict[str, Dict[str, Any]] = {}

@app.post("/submit-letter")
async def submit_letter(
    text: str = Form(...),
    addressee: str = Form(""),
    salutation: str = Form(...),
    date: str = Form(...)
):
    """
    Receive letter data via POST and redirect to Streamlit app with letter ID
    """
    # Generate unique ID for this letter
    letter_id = str(uuid.uuid4())[:8]
    
    # Store the letter data with timestamp for cleanup
    letter_storage[letter_id] = {
        "text": text,
        "addressee": addressee,
        "salutation": salutation,
        "date": date,
        "timestamp": time.time()
    }
    
    # Redirect to Streamlit app with the letter ID
    return RedirectResponse(
        url=f"https://letter-crafter.streamlit.app/?letter_id={letter_id}",
        status_code=302
    )

@app.get("/get-letter/{letter_id}")
async def get_letter(letter_id: str):
    """
    Retrieve stored letter data by ID
    """
    if letter_id not in letter_storage:
        raise HTTPException(status_code=404, detail="Letter not found")
    
    return letter_storage[letter_id]

@app.get("/health")
async def health_check():
    return {"status": "healthy"}

# Cleanup old letters (optional background task)
@app.on_event("startup")
async def cleanup_old_letters():
    import asyncio
    
    async def cleanup():
        while True:
            current_time = time.time()
            expired_ids = [
                letter_id for letter_id, data in letter_storage.items()
                if current_time - data["timestamp"] > 3600  # 1 hour
            ]
            for letter_id in expired_ids:
                del letter_storage[letter_id]
            await asyncio.sleep(300)  # Clean up every 5 minutes
    
    asyncio.create_task(cleanup())

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
