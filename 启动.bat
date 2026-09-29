@echo off
cd /D D:\github\novel_rag\novel
start http://localhost:8501
python -m streamlit run app.py --server.headless true
pause
