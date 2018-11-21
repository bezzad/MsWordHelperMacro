@echo off

copy "./Normal.dotm" "C:\Users\%username%\AppData\Roaming\Microsoft\Templates"
echo "***Macros updated successfuly"
copy "./Word.officeUI" "C:\Users\%username%\AppData\Local\Microsoft\Office"
echo "***Riboons updated successfuly"


set /P = "Press any key ..."