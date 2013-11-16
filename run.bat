@echo off
START /wait dist/mhsCreater.exe || goto QUIT

:QUIT
	PAUSE