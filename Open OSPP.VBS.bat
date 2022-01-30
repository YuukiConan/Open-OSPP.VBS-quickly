@echo off

echo ---------This Batch Created Using BatchFile Language----------

IF DEFINED (%ProgramFiles(x86)) (
goto loading

:loading

echo Please Wait...
set TARGET_PATH= %ProgramFiles(x86)%
goto openapp

:openapp

echo Open OSPP.VBS....
start "%TARGET_PATH%\Microsoft Office\Office16\OSPP.VBS
goto load1

:load1

       set path= "%TARGET_PATH%\Microsoft Office\Office16\"
     set file= "%TARGET_PATH%\Microsoft Office\Office16\SLERROR.xml
   copy= "%TARGET_PATH%\Microsoft Office\Office16\SLERROR.xml"
  copy to= "C://Users/Admin/Dekstop/"
 goto load2

:load2

 set file= "%TARGET_PATH%\Microsoft Office\Office16\SLERROR.xml"
  set file= "%TARGET_PATH%\Microsoft Office\Office16\vNextDiag.ps1"
   set to= "%TARGET_PATH%\Microsoft Office\Office16\OSPP.VBS"
    goto end

:end

echo Done!
pause
