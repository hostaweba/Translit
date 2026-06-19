call translit\scripts\activate
python main.py


echo. 
echo.
echo -------------------------------------
echo Look above. Is everything OK?
echo -------------------------------------
pause >nul

cls


call CustomKeys.bat

echo.
echo -------------------------------------
dir | find ".py"
echo -------------------------------------
echo.
echo -------------------------------------
dir | find ".bat"
echo -------------------------------------
echo. 
echo.

CMD  