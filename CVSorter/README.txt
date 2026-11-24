========================================================
   CV SORTER - Teacher CVs Auto Sorting Desktop App
========================================================

YE APP KYA KARTI HAI
--------------------
- Aap CV add karte hain (PDF, Word .docx, ya .txt)
- App CV parh kar khud detect karti hai ke ye kis subject ke teacher ki hai
- Har subject ka alag folder bana kar CV usme save kar deti hai, jaise:
     Documents\Sorted CVs\English Teacher\
     Documents\Sorted CVs\Math Teacher\
     Documents\Sorted CVs\Physics Teacher\  ... waghera

BUILT-IN SUBJECTS (11)
----------------------
English, Math, Physics, Chemistry, Biology, Computer Science, Urdu,
Islamiat, General Science, Social Studies/Pak Studies, Primary/Montessori.
Settings > "Subjects & Keywords" se naye subject (jaise Arabic Teacher) add kar sakte hain.

EXTRA FEATURES
--------------
1. Ek saath bohat saari CVs ya poora folder add karein
2. Drag & Drop (files ko window mein ghaseet kar chhor dein)
3. Candidate ka Naam, Email, Phone, Experience khud nikalti hai
4. File ka naam "Ayesha Khan - English Teacher.pdf" kar deti hai
5. Duplicate CV pakadti hai (same CV do baar save nahi hogi)
6. Jo CV samajh na aaye woh "Unsorted - Manual Review" folder mein jati hai
7. Galat detect ho to row par Right-Click > Change Subject (file bhi move ho jayegi)
8. Search box: naam, subject, email ya phone se dhundein
9. Excel Report: saari CVs ki list + subject-wise summary
10. Auto-Watch: aik folder (jaise Downloads ya email attachments) set karein,
    jo bhi nayi CV aaye app khud sort kar degi
11. Copy ya Move mode, aur match strictness Settings se badal sakte hain
12. Har CV ka record save rehta hai, app band karke kholne par bhi

PEHLI DAFA CHALANE KA TAREEQA (Windows)
---------------------------------------
1. Python install karein: https://www.python.org/downloads/
   (Install ke waqt "Add Python to PATH" ka box zaroor tick karein)
2. Ye folder kisi jagah unzip karein
3. "1_install.bat" par double-click (sirf pehli dafa, internet chahiye)
4. "2_run_app.bat" par double-click  -> app khul jayegi

.EXE BANANA (taake bina Python ke kisi bhi PC par chale)
-------------------------------------------------------
- Step 3 ke baad "3_build_exe.bat" chalayein
- "dist" folder mein "CV Sorter.exe" ban jayegi
- Isko desktop par rakh kar shortcut ki tarah use karein

COMPANY KA NAAM LAGANA
----------------------
cv_sorter_app.py ko Notepad mein kholein, ye line dhundein:
     COMPANY_NAME = "Your Company Name"
aur apni company ka naam likh dein. Phir dobara .exe banayein.

ZAROORI BAATEIN
---------------
- Scanned (photo wali) PDF ka text nahi parha ja sakta, woh "Manual Review" mein jayegi
- Purani .doc files support nahi; unhein Word mein khol kar .docx save karein
- Settings aur database yahan save hote hain:  C:\Users\<aapka naam>\CVSorter\
