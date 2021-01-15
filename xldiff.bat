@ECHO OFF

REM Create paths to original and current spreadsheets to store in tmp
set path1=%2
set path2=%5

REM Change forward slash to back slash on all paths for the Excel tool
set path1=%path1:/=\%
set path2=%path2:/=\%

REM Write path to a file, so SpreadSheetCompare could use it as one parameter
ECHO %path1% > tmp.txt
ECHO %path2% >> tmp.txt

REM Set Compare Program, this might vary per system
set prog="C:\Program Files\Microsoft Office\root\Client\AppVLP.exe" "C:\Program Files (x86)\Microsoft Office\Office16\DCF\SPREADSHEETCOMPARE.EXE"

REM Run Compare
REM in %appdata% Local/Temp
%prog% tmp.txt
REM pause to provent system from removing git temp checkout
ECHO Begin run xldiff.bat to launch SpreadsheetCompare ...
ECHO bat will pause here, when complete, press enter to continue.
pause
