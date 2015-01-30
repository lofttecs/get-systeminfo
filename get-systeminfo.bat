@echo off

pushd %~dp0

FOR /F "usebackq" %%t IN (`HOSTNAME`) DO SET HNAME=%%t
SET TODAY=%date:~-10,4%%date:~-5,2%%date:~-2,2%
SET FILE_O=log\%HNAME%__%TODAY%.txt
SET FILE_T=log\%HNAME%_tmp.txt

"%ProgramFiles%\Common Files\Microsoft Shared\MSInfo\msinfo32" /report %FILE_O%

cmd /u /c echo; >> %FILE_O%
cmd /u /c echo --------"wmic csproduct get"-------- >> %FILE_O%
cmd /u /c echo; >> %FILE_O%
wmic csproduct get /value >> %FILE_O%

cmd /u /c echo; >> %FILE_O%
cmd /u /c echo --------"wmic memorychip get"-------- >> %FILE_O%
cmd /u /c echo; >> %FILE_O%
wmic memorychip get /value >> %FILE_O%

cmd /u /c echo; >> %FILE_O%
cmd /u /c echo --------"wmic memlogical get"-------- >> %FILE_O%
cmd /u /c echo; >> %FILE_O%
wmic memlogical get /value >> %FILE_O%

cmd /u /c echo; >> %FILE_O%
cmd /u /c echo --------"wmic product get"-------- >> %FILE_O%
cmd /u /c echo; >> %FILE_O%
wmic product get name /value >> %FILE_O%

cmd /u /c echo; >> %FILE_O%
cmd /u /c echo --------"systeminfo"-------- >> %FILE_O%
cmd /u /c echo; >> %FILE_O%
systeminfo /fo list > %FILE_T%
cmd /u /c type %FILE_T% >> %FILE_O%
del %FILE_T%

net use /delete %~dp0

echo WScript.Echo "complete" > %TEMP%\msgboxtest.vbs & %TEMP%\msgboxtest.vbs
del %TEMP%\msgboxtest.vbs
