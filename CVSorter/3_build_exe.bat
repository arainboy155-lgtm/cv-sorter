@echo off
echo CV Sorter ki .exe file ban rahi hai (2-3 minute lagenge)...
python -m PyInstaller --noconfirm --onefile --windowed --name "CV Sorter" --collect-all tkinterdnd2 --collect-all pdfplumber cv_sorter_app.py
echo.
echo Mukammal! Aapki app yahan hai:  dist\CV Sorter.exe
echo Is .exe ko kisi bhi Windows PC par copy karke chala sakte hain (Python ki zaroorat nahi).
pause
