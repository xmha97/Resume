@echo off
del "%temp%\Document.zip"
del "%temp%\Document.pptx"
powershell Compress-Archive -Path "%~dp0Document\*" -DestinationPath "%temp%\Document.zip"
ren "%temp%\Document.zip" "Document.pptx"
start /wait "" "C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE" "%temp%\Document.pptx"
rmdir /S /Q "Document"
ren "%temp%\Document.pptx" "Document.zip"
powershell expand-archive "%temp%\Document.zip" "%~dp0Document"
