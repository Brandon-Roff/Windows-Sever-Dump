@echo off
color 02
echo. 
echo 			........ USB Copy ........
echo.
set /p "		mysource=Enter Source Folder: "
echo 
set /p "		mydest=Enter destination Folder: "
xcopy "%mysource%" "%mydest%" /f /-Y /E

sleep 5
