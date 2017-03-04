@echo off
python -m pip install --upgrade pip
echo.
for %%s in (beautifulsoup4 requests xlsxwriter) do (
    pip install %%s
    echo.
) 
pause