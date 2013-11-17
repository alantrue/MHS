@echo off
START /wait v64/dist/mhsCreater.exe || goto QUIT

:QUIT
	PAUSE