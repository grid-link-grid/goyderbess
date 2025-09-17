@echo off
echo Deleting .aux, .log, and .gz files in the current directory...

del *.aux 2>nul
del *.log 2>nul
del *.gz 2>nul

echo Deletion complete.