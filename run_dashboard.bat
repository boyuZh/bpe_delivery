@echo off
echo Installing dependencies...
@REM pip install -r requirements.txt

echo Starting dashboard application...
python -m streamlit run dashboard.py

pause