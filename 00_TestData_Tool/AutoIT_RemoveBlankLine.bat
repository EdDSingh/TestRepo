echo %1
set "MyPath=%~dpnx0" & call set "MyPath=%%MyPath:\%~nx0=%%" 
cd %MyPath% 
RemoveEmptyLines.exe %1