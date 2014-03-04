@echo off
cd /d %~dp0

Wscript file_distributer.vbs M .\move_lists\move.txt .\move_test

pause

exit
