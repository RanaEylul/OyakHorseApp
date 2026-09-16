@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

set "PY_EXE="

rem WinPython klasorunu ve icindeki python.exe dosyasini bulur
for /d %%G in (WPy*) do (
    for /d %%H in ("%%G\python*") do (
        if exist "%%~fH\python.exe" (
            set "PY_EXE=%%~fH\python.exe"
        )
    )
)

rem Eger WinPython bulunamazsa bilgisayardaki Python'a fallback yapar
if not defined PY_EXE (
    set "PY_EXE=C:\Users\ranae\AppData\Local\Python\bin\python.exe"
)

rem Streamlit uygulamasini baslatir
"!PY_EXE!" -m streamlit run fabrikaH_app.py
pause