@echo off

:ADMIN
openfiles >nul 2>nul ||(
  echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
  echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"
  "%temp%\getadmin.vbs" >nul 2>&1
  goto:eof
)
del /f /q "%temp%\getadmin.vbs" >nul 2>nul
pushd "%~dp0"

cd /d %ProgramFiles%\Microsoft Office\Office16
cd /d %ProgramFiles(x86)%\Microsoft Office\Office16

cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectStdVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectStdVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectStdVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectStdVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectStdVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectStdVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectStdVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inpkey:GNFHQ-F6YQM-KQDGJ-327XX-KQBVC
cscript ospp.vbs /sethst:www.ddddg.cn
cscript ospp.vbs /act
