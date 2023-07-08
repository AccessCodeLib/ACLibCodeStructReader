@echo off

if exist .\ACLibStructReader.accdb (
set /p CopyFile=ACLibStructReader.accdb exists .. overwrite with access-add-in\ACLibStructReader.accda? [Y/N]:
) else (
set CopyFile=Y
)

if /I %CopyFile% == Y (
	echo File is copied ...
) else (
	echo Batch is cancelled
	pause
	exit
)

copy .\access-add-in\ACLibStructReader.accda ACLibStructReader.accdb

timeout 2