@REM Maven Wrapper script for Windows
@echo off
setlocal

set "MAVEN_PROJECTBASEDIR=%~dp0"
set "MAVEN_WRAPPER_PROPERTIES=%MAVEN_PROJECTBASEDIR%.mvn\wrapper\maven-wrapper.properties"

if exist "%MAVEN_WRAPPER_PROPERTIES%" (
    for /f "tokens=2 delims==" %%a in ('findstr "distributionUrl" "%MAVEN_WRAPPER_PROPERTIES%"') do set "DISTRIBUTION_URL=%%a"
)

mvn %*
