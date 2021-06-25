# escape=`

FROM mcr.microsoft.com/windows:1903
SHELL ["powershell", "-ExecutionPolicy Bypass", "-Command", "$ErrorActionPreference = 'Stop'; $ProgressPreference = 'SilentlyContinue';"]
COPY . C:\build
WORKDIR C:\build
RUN .\vb6setup.ps1