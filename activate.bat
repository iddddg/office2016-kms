@echo off
rem 淡定定定哥 https://www.ddddg.cn 自用.版权就别删除了吧,感激
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

cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProPlusVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ProjectProVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\VisioProVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99
cscript ospp.vbs /inpkey:YG9NW-3K39V-2T3HJ-93F3Q-G83KT
cscript ospp.vbs /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK

cscript ospp.vbs /sethst:kms.ddddg.cn
cscript ospp.vbs /act
