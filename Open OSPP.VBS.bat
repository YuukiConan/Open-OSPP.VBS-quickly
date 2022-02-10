@echo off

echo ---------This Batch Created Using BatchFile Language----------
echo Working with Office 2016 and Windows x64 System

IF DEFINED (%ProgramFiles(x86)%) (
goto loading
)

:loading

title = "######your name project" ######optional
echo Please Wait...
set TARGET_PATH= %ProgramFiles(x86)%
goto openapp

:openapp

echo Open OSPP.VBS....
start "%TARGET_PATH%\Microsoft Office\Office16\OSPP.VBS"
goto load

 :load

     echo Copying OSPP.VBS....
        set path= "%TARGET_PATH%\Microsoft Office\Office16"
        set file= %TARGET_PATH%\Microsoft Office\Office16\OSPP.VBS"
        copy= "%TARGET_PATH%\Microsoft Office\Office16\SLERROR.xml"
        copy to= "C://Users/Admin/Dekstop/"
 goto load1

:load1

     echo Copying SLERROR.XML....
       set path= "%TARGET_PATH%\Microsoft Office\Office16"
       set file= "%TARGET_PATH%\Microsoft Office\Office16\SLERROR.xml"
       copy= "%TARGET_PATH%\Microsoft Office\Office16\SLERROR.xml"
       copy to= "C://Users/Admin/Dekstop/"
 goto load2

:load2

       set file= "%TARGET_PATH%\Microsoft Office\Office16\SLERROR.xml"
       set file= "%TARGET_PATH%\Microsoft Office\Office16\OSPP.HTM"
       set file= "%TARGET_PATH%\Microsoft Office\Office16\vNextDiag.ps1"
    set to= "%TARGET_PATH%\Microsoft Office\Office16\OSPP.VBS"

:end

echo Done!
pause
