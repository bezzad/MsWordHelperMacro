@echo off

copy "C:\Users\%username%\AppData\Roaming\Microsoft\Templates\Normal.dotm" ".\macros_installer\Normal.dotm"
echo "***Macros updated successfuly"
copy "C:\Users\%username%\AppData\Local\Microsoft\Office\Word.officeUI" ".\macros_installer\Word.officeUI"
echo "***Riboons updated successfuly"


set /P = "Press any key ..."