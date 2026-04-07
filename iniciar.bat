@echo off
echo ============================================
echo   App Dotacion SSMC - Iniciando
echo ============================================
echo.
set PYTHON_EXE=C:\Users\DAP\AppData\Local\Programs\Python\Python314\python.exe
echo Verificando dependencias...
"%PYTHON_EXE%" -m pip install -r requirements.txt --quiet 2>nul
echo.
echo Iniciando aplicacion...
"%PYTHON_EXE%" -m streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
pause
