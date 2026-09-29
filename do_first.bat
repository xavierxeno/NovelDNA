@echo off
chcp 65001 >nul
cd /d "%~dp0"
title 安装 Python 依赖

echo ======================================
echo   作品语料检索与风格解析器 - 依赖安装
echo ======================================
echo.

echo [1/3] 检查 Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 没检测到 Python，请先安装 Python 3.10+
    pause
    exit /b
)
python --version

echo.
echo [2/3] 升级 pip...
python -m pip install --upgrade pip

echo.
echo [3/3] 安装依赖...
pip install streamlit chromadb ollama openai python-docx pywin32

echo.
echo ======================================
echo   ✅ 依赖安装完成！
echo   现在可以双击 启动.bat
echo ======================================
pause
