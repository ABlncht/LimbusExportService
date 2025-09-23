xcopy /s /y "C:\Users\utilisateur\source\repos\LimbusExportService\bin\Release\*" "C:\ProgramData\LimbusExportService\"
cd C:\Windows\Microsoft.NET\Framework64\v4.0.30319
InstallUtil.exe "C:\ProgramData\LimbusExportService\LimbusExportService.exe"
@echo off
echo Installation OK, appuyez sur une touche pour continuer…
pause > nul

