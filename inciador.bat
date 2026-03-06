@echo off
cd backend
call venv\Scripts\activate
python -m uvicorn app.main:app --reload
pause