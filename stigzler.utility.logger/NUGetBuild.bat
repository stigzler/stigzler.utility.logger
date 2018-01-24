@echo off
setlocal enableextensions enabledelayedexpansion
:: NUGetBuilder - by stigzler
:: Put this .bat file in the project root folder with the .vbproj or .csproj File
:: Also download and place nuget.exe in the same folder
:: Download here: https://www.nuget.org/downloads
::
:: Visual Studio Setup
:: Select Case bat file In the solution explorer And choose File...Save myfile.bat As...
:: On the down arrow on the Save button, choose Save with Encoding.
:: When saving in the Advanced Save Options dialog, Choose US-ASCII in the Encoding drop-down list. 
:: Set the line endings As required, Or leave it As Current Setting.

:: Uncomment out below to stop cmd window closing on error:
:: if not defined in_subprocess (cmd /k set in_subprocess=y ^& %0 %*) & exit )

:: USER VARIABLES - SET THESE
Set App=stigzler.utility.logger
Set APIKey=oy2cg7gahmir2d4rk3ya7j65ydwykj3ene6dhn3wvt6tje

:: CODE
cd /d %~dp0

echo %App% NuGet Deploy
echo Info: https://docs.microsoft.com/en-us/nuget/quickstart/create-and-publish-a-package
echo(

echo Checking .nuspec file exists..
if not exist %app%.nuspec (
	ECHO %app%.nuspec doesn't exist. Creating..
	nuget spec -Force
	echo Now edit %app%.nuspec and then press any key..
	%App%.nuspec
	pause
	) else (
	echo %app%.nuspec already exists.
	)

echo(
echo Getting Verison from XML file..
    set "version="
    for /f "tokens=3 delims=<>" %%a in (
        'find /i "<version>" ^< "%app%.nuspec"'
    ) do set "version=%%a"
echo Version: [%version%]

echo(
echo Getting ID from XML file..
    set "ID="
    for /f "tokens=3 delims=<>" %%a in (
        'find /i "<ID>" ^< "%app%.nuspec"'
    ) do set "ID=%%a"
echo ID: [%ID%]

echo(
echo Building nuget package
nuget pack %App%.vbproj

echo(
echo API Key...
if "%APIKey%" == "" (
	echo No API Key set.
	set /p APIKey=Enter the API key from NuGet: 
)

echo APIKey set to: %APIKey%

echo(
echo Publishing Project..
nuget push %ID%.%version%.nupkg %APIKey% -Source https://api.nuget.org/v3/index.json

echo(
echo Done. Any key to end.
echo(
pause
goto :eof

::SUBROUTINES ==================================================

:EOF

:: EXAMPLE .nuspec FILE:
REM <?xml version="1.0"?>
REM <package>
  REM <metadata>
    REM <id>stigzler.utility.updater</id>
    REM <version>1.0.0</version>
    REM <title>StigzlerUpdater</title>
    REM <authors>stigzler</authors>
    REM <owners>stigzler</owners>
    REM <requireLicenseAcceptance>false</requireLicenseAcceptance>
    REM <description>Update your applications via http of zip file..</description>
    REM <releaseNotes>Initial release.</releaseNotes>
    REM <copyright>Copyright 2018</copyright>
    REM <tags>update</tags>
  REM </metadata>
REM </package>
