@echo off
title Goose's Hyper-V VM Connect Tool

rem =========================================
rem ======== Hyper-V VM Connect Tool ========
rem == Created by Splendead Goose 03.19.19 ==
rem =========================================

rem ====================================================
rem == Change Log ======================================
rem == Mar 19, 2019 - Initial Script Creation ==========
rem == Mar 20, 2019 - Working Script to get Local VMs ==
rem == Mar 21, 2019 - Option to Restart Service ========
rem == Mar 22, 2019 - Cluster Options in Progress ======
rem == Mar 25, 2019 - Cluster Options Complete =========
rem == Mar 26, 2019 - Cluster VMs No Longer Truncated ==
rem == May 28, 2023 - Minor Tweaks for Public Release ==
rem ====================================================

rem ========================
rem === Obtain the ADMIN ===
rem ========================

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "%~s0", "%params%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------

rem =================================
rem === Get localhost Running VMs ===
rem =================================

:InitialVMList
set /a VMList=1
set /a EDIT=0
set /a MODE=0
set /a CLST=0
set /a CONNECTCLST=0
set HOST=%COMPUTERNAME%
set DOMAINS=%USERDNSDOMAIN%
set File2Read=%temp%\vms.txt
call:AllVMs

rem =====================================
rem === List VMs Switches and Options ===
rem =====================================

:ListOptions
title Goose's Hyper-V VM Connect Tool
rem if not exist %File2Read% (goto :Error)
cls
set /a NUMBER=0
set /a COUNT=0
color F
echo Welcome to Goose's Hyper-V VM Connect Tool!
echo.
if /i {%CLST%}=={0} (echo VMs on: %HOST%)
if /i {%CLST%}=={1} (echo Clusters on: %DOMAINS%)
echo.
setlocal EnableDelayedExpansion
for /f "tokens=1,3 delims=" %%a in ('type "%File2Read%"') do (
	set /a COUNT+=1
	set Line[!COUNT!]=%%a
	set VMN[!COUNT!]==%%b
)
for /L %%i in (1,1,%COUNT%) do (
	echo 	%%i. !Line[%%i]!
)
echo.
echo Switches:
echo.
if /i {%VMList%}=={0} (echo 	D. Display VM State - Running)
if /i {%VMList%}=={1} (echo 	D. Display VM State - All VMs)
if /i {%EDIT%}=={0} (echo 	E. Edit Connection - OFF)
if /i {%EDIT%}=={1} (echo 	E. Edit Connection - ON)
if /i {%MODE%}=={0} (echo 	M. Connection Mode - Console)
if /i {%MODE%}=={1} (echo 	M. Connection Mode - RDP)
echo.
echo Options:
echo.
echo 	C. Custom Connection
echo 	G. Get VMs on Another Host
echo 	P. Get Clusters on %DOMAINS%
echo 	Q. Get VMs on Cluster
echo 	R. Restart Hyper-V Service
echo 	Z. Reset List
echo.
echo 	0. Exit Goose's Hyper-V VM Connect Tool
echo.
set /P NUMBER=Enter Number or Letter:  
echo.
if /i {%NUMBER%}=={0} (goto :0)
if /i {%NUMBER%}=={C} (call:StartConsole)
if /i {%NUMBER%}=={D} (call:DisplaySwitch)
if /i {%NUMBER%}=={E} (call:EditSwitch)
if /i {%NUMBER%}=={G} (call:GetVMsFromHost)
if /i {%NUMBER%}=={M} (call:ConnectionMode)
if /i {%NUMBER%}=={P} (call:Clusters)
if /i {%NUMBER%}=={Q} (
	set /p CLUSTNAME=Please Enter Cluster Name: 
	call:ClusterVMs !CLUSTNAME!
)
if /i {%NUMBER%}=={R} (call:RestartService)
if not {!Line[%NUMBER%]!}=={} (
	if /i {%CLST%}=={0} (call:StartConsole "!Line[%NUMBER%]!")
	if /i {%CLST%}=={1} (call:ClusterVMs !Line[%NUMBER%]!)
)
if /i {%NUMBER%}=={Z} (call:InitialVMList)
endlocal
goto :ListOptions

rem =================================
rem ==== FUNCTIONS... Well Kinda ====
rem =================================

rem ====================================
rem === Start the Console Connection ===
rem ====================================

:StartConsole
set VMNAME=%~1
set CLSTVMNAME=%~2
if /i {%CONNECTCLST%}=={1} (
	start vmconnect %VMNAME%
	goto :EOF
)
if /i {%MODE%}=={0} (
	if /i {%NUMBER%}=={C} (
		set /p HOST=Please Enter Server Hostname: 
		set /p VMNAME=Please Enter VM Name: 
	)
	if /i {%EDIT%}=={0} (start vmconnect %HOST% "%VMNAME%")
	if /i {%EDIT%}=={1} (start vmconnect %HOST% "%VMNAME%" /edit)
)
if /i {%MODE%}=={1} (
	if /i {%NUMBER%}=={C} (
		set /p VMNAME=Please Enter Computer Name: 
	)
	mstsc /v:%VMNAME%
)

goto :EOF

rem =========================================
rem === Switch Between Running or All VMs ===
rem =========================================

:DisplaySwitch
if /i {%VMList%}=={0} (
	call:AllVMs
	set /a VMList=1
	goto :ListOptions
)
if /i {%VMList%}=={1} (
	call:RunningVMs
	set /a VMList=0
	goto :ListOptions
)
goto :EOF

rem ======================================
rem === Edit Enhanced Session Settings ===
rem ======================================

:EditSwitch
if /i {%EDIT%}=={0} (
	set /a EDIT=1
	goto :ListOptions
)
if /i {%EDIT%}=={1} (
	set /a EDIT=0
	goto :ListOptions
)
goto :EOF

rem ==============================
rem === Switch Connection Mode ===
rem ==============================

:ConnectionMode
if /i {%MODE%}=={0} (
	call:GetVMCompName
	set /a MODE=1
	set /a VMList=1
	goto :ListOptions
)
if /i {%MODE%}=={1} (
	if /i {%VMList%}=={0} (call:RunningVMs)
	if /i {%VMList%}=={1} (call:AllVMs)
	set /a MODE=0
	goto :ListOptions
)
goto :EOF

rem ====================================
rem === Get VMs from a Specific Host ===
rem ====================================

:GetVMsFromHost
set /p HOST=Please Enter Server Hostname: 
echo.
call:AllVMs
set /a VMList=1
set /a CLST=0
set /a CONNECTCLST=0
goto :ListOptions

rem ===================
rem === Get All VMs ===
rem ===================

:AllVMs
echo Obtaining All VMs on %HOST%...
powershell -executionpolicy bypass "Get-VM -ComputerName %HOST% | foreach { $_.Name } > $env:temp\vms.txt"
goto :EOF

rem =======================
rem === Get Running VMs ===
rem =======================

:RunningVMs
echo Obtaining Running VMs on %HOST%...
powershell -executionpolicy bypass "Get-VM -ComputerName %HOST% | Where { $_.State -eq 'Running' } | foreach { $_.Name } > $env:temp\vms.txt"
goto :EOF

rem =======================
rem === Get IP from VMs ===
rem =======================

:GetVMCompName
echo Obtaining VM Hostnames on %HOST%...
rem powershell -executionpolicy bypass "((Get-VM -ComputerName %HOST% | Where { $_.State -eq 'Running' }) | Select -ExpandProperty NetworkAdapters).ipaddresses[0] | ForEach-Object {([System.Net.Dns]::GetHostEntry($_)).Hostname } > $env:temp\vms.txt"
powershell -executionpolicy bypass "((Get-VM -ComputerName %HOST% | Where { $_.State -eq 'Running' }) | Select -ExpandProperty NetworkAdapters).ipaddresses[0] > $env:temp\vms.txt"
goto :EOF

rem ===============================
rem === Restart Hyper-V Service ===
rem ===============================

:RestartService
echo PLEASE MAKE SURE NO VM IS RUNNING
timeout /t 99999 > NUL
echo.
echo Stopping VMMS Service...
sc stop vmms
timeout /t 5 > NUL
echo.
echo Starting VMMS Service...
sc start vmms
timeout /t 5 > NUL
if /i {%VMList%}=={0} (call:RunningVMs)
if /i {%VMList%}=={1} (call:AllVMs)
goto :ListOptions

rem ====================
rem === Get Clusters ===
rem ====================

:Clusters
echo Obtaining Clusters on %DOMAINS%...
powershell -executionpolicy bypass "Get-Cluster -Domain %DOMAINS% | foreach { $_.Name } > $env:temp\vms.txt"
set /a VMList=2
set /a EDIT=2
set /a CLST=1
set /a CONNECTCLST=0
goto :ListOptions

rem ==========================
rem === Get VMs on Cluster ===
rem ==========================

:ClusterVMs
set HOST=%~1
echo.
echo Obtaining VMs on %HOST%
rem powershell -executionpolicy bypass "Get-VM -ComputerName (Get-ClusterNode -Cluster %HOST%) | foreach { $_.Name } > $env:temp\vms.txt"
powershell -executionpolicy bypass "Get-VM -ComputerName (Get-ClusterNode -Cluster %HOST%) | select -Property ComputerName, Name | Format-Table -AutoSize -HideTableHeaders | Out-File -Width 512 $env:temp\vms.txt"
set /a VMList=2
set /a EDIT=0
set /a CLST=0
set /a CONNECTCLST=1
goto :ListOptions

rem =======================================
rem === Error Catch for Reading VM List ===
rem =======================================

:Error
echo Error Reading %File2Read%
timeout /t 99999 > NUL

:0
del %File2Read%
exit