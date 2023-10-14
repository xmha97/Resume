@echo off
del "%temp%\Resume.zip"
del "%temp%\Resume.pptx"
powershell Compress-Archive -Path "%~dp0Resume\*" -DestinationPath "%temp%\Resume.zip"
ren "%temp%\Resume.zip" "Resume.pptx"
start /wait "" "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" "%temp%\Resume.pptx"
rmdir /S /Q "Resume"
ren "%temp%\Resume.pptx" "Resume.zip"
powershell expand-archive "%temp%\Resume.zip" "%~dp0Resume"
