@echo off
REM Ir a la carpeta donde está este .bat
cd /d "%~dp0"

REM Activar entorno virtual si tienes uno (OPCIONAL)
REM call venv\Scripts\activate

REM Ejecutar la app de Streamlit
python -m streamlit run app.py

REM Mantener la ventana abierta al final (útil para ver errores)
pause
