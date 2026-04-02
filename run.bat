@echo off
echo [1/3] Checking dependencies...
:: Установка библиотек из requirements.txt
python -m pip install -r requirements.txt

echo.
echo [2/3] Starting Streamlit app...
:: Запуск через модуль python -m, чтобы обойти ошибку с PATH
python -m streamlit run main.py

echo.
echo [3/3] Application closed.
pause
