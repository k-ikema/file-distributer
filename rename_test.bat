@echo off
cd /d %~dp0

Wscript file_distributer.vbs R .\rename_lists\rename.txt .\rename_test

pause

exit
