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

cscript ospp.vbs /inslic:"..\root\Licenses16\ExcelVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ExcelVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ExcelVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ExcelVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ExcelVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ExcelVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\ExcelVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inslic:"..\root\Licenses16\OutlookVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\OutlookVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\OutlookVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\OutlookVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\OutlookVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\OutlookVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\OutlookVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inslic:"..\root\Licenses16\PowerPointVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\PowerPointVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\PowerPointVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\PowerPointVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\PowerPointVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\PowerPointVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\PowerPointVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inslic:"..\root\Licenses16\WordVL_KMS_Client-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\WordVL_KMS_Client-ul.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\WordVL_KMS_Client-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\WordVL_MAK-pl.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\WordVL_MAK-ppd.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\WordVL_MAK-ul-oob.xrm-ms"
cscript ospp.vbs /inslic:"..\root\Licenses16\WordVL_MAK-ul-phn.xrm-ms"

cscript ospp.vbs /inpkey:9C2PK-NWTVB-JMPW8-BFT28-7FTBF
cscript ospp.vbs /inpkey:R69KK-NTPKF-7M3Q4-QYBHW-6MT9B
cscript ospp.vbs /inpkey:J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6
cscript ospp.vbs /inpkey:WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6

cscript ospp.vbs /sethst:www.ddddg.cn
cscript ospp.vbs /act
