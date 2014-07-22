:: Purpose:		Runs a series of tools and checks on the system to help automate
::				our Monthly Server Maintenance
:: Author:      Jason Cobb ITCC
:: Version:     2.0 See notes on version 1.4 to see changes
::
:: Usage:       Run this script and then manually review the log files created

:: Log notes, Make sure to write to log and screen so that it is both shown and logged
:: echo INFONEEDEDTOWRITETOLOG >> LOGFILE
:: echo INFONEEDEDTOWRITETOLOG 

:: Pre Run setup
@echo off && cls
color fc
::Create Variables and log folder\file
SETLOCAL
set VERSION=2.0
set UPDATED=2014-07-10
set CUR_DATE=%DATE:~-4%-%DATE:~4,2%-%DATE:~7,2%
set LOGPATH=%SystemDrive%\Temp\MSM\Logs-%CUR_DATE%
set LOGFILE=MSM.%CUR_DATE%.log

if not exist %LOGPATH% mkdir %LOGPATH%
if not exist %LOGPATH%\%LOGFILE% echo. > %LOGPATH%\%LOGFILE%
title MSMScript v%VERSION% (Updated on %UPDATED%)


::::::::::::::::::::
:: WELCOME SCREEN ::
::::::::::::::::::::
:welcome_screen
echo ************************************************************* > %LOGPATH%\%LOGFILE%
echo *************************************************************
echo * MSMScript v%VERSION%  ( Created %UPDATED%)                     * >> %LOGPATH%\%LOGFILE%
echo * MSMScript v%VERSION%  ( Created %UPDATED%)                     *
echo * Author: Jason Cobb, Systems Engineer, ITCC                * >> %LOGPATH%\%LOGFILE%
echo * Author: Jason Cobb, Systems Engineer, ITCC                *
echo * Date: %Date% Time: %TIME%                    * >> %LOGPATH%\%LOGFILE%
echo * Date: %Date% Time: %TIME%                    *
echo * --------------------------------------------------------- * >> %LOGPATH%\%LOGFILE%
echo * --------------------------------------------------------- *
echo *  Once this completes check the log files created to verify* >> %LOGPATH%\%LOGFILE%
echo *  Once this completes check the log files created to verify*
echo *  That everything is up to date and created correctly      * >> %LOGPATH%\%LOGFILE%
echo *  That everything is up to date and created correctly      *
echo ************************************************************* >> %LOGPATH%\%LOGFILE%
echo *************************************************************
echo %LOGPATH%

:::::::::::::::::::::
::stage_diskspace  ::
::Check free space ::
:::::::::::::::::::::
:stage_diskspace
title MSMScript v%VERSION% Checking Disk Space
echo %CUR_DATE% %TIME% Starting Disk Space Check >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Starting Disk Space Check
echo %CUR_DATE% %TIME% Running driveSpace.vbs >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Running driveSpace.vbs 
pushd resources
cscript driveSpace.vbs %LOGPATH%
echo ------------------------------------------ >> %LOGPATH%\%LOGFILE%
echo ------------------------------------------
popd

::::::::::::::::::
::stage_chkdsk  ::
::Run CHKDSK on ::
::All drives    ::
::::::::::::::::::
:stage_chkdsk
title MSMScript v%VERSION% Running CHKDSK
echo %CUR_DATE% %TIME% Starting CHKDSK On each drive >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Starting CHKDSK On each drive
echo %CUR_DATE% %TIME% Running chkdskDrives.vbs >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Running chkdskDrives.vbs
pushd resources
cscript chkdskDrives.vbs %LOGPATH%
echo ------------------------------------------ >> %LOGPATH%\%LOGFILE%
echo ------------------------------------------
popd

::::::::::::::::::::
::stage_biosinfo  ::
::Run SystemInfo  ::
::To get BIOS     ::
::::::::::::::::::::
:stage_biosinfo
title MSMScript v%VERSION% Running SYSTEMINFO
echo %CUR_DATE% %TIME% Running SystemInfo >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Running SystemInfo
pushd resources
cscript SystemInfoGather.vbs %LOGPATH%
echo ------------------------------------------ >> %LOGPATH%\%LOGFILE%
echo ------------------------------------------
popd

:::::::::::::::::
::stage_netstat::
::Run netstat  ::
:::::::::::::::::
:stage_netstat
title MSMScript v%VERSION% Running Netstat -an
echo %CUR_DATE% %TIME% Running netstat -an >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Running netstat -an
netstat -an > %LOGPATH%\Netstat.%CUR_DATE%.log
echo ------------------------------------------ >> %LOGPATH%\%LOGFILE%
echo ------------------------------------------

::::::::::::::::::::::::::::::::::
:: If adding additional modules ::
:: make sure to add them before ::
:: The stage_parse section      ::
:: and to modify the            ::
:: ParseData.vbs file if info   ::
:: needs to be included in the  ::
:: final report.                ::
::::::::::::::::::::::::::::::::::



:::::::::::::::::
::stage_parse  ::
::Parse Data   ::
:::::::::::::::::
:stage_parse
title MSMScript v%VERSION% Parsing Collected Data
echo %CUR_DATE% %TIME% Starting to parse the data collected >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Starting to parse the data collected
echo %CUR_DATE% %TIME% Running ParseData.vbs >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% Running ParseData.vbs
pushd resources
cscript ParseData.vbs %LOGPATH% %LOGFILE%
popd
echo ------------------------------------------ >> %LOGPATH%\%LOGFILE%
echo ------------------------------------------

:::::::::::::
::Completed::
:::::::::::::
:completed
title MSMScript v%VERSION% COMPLETED
echo %CUR_DATE% %TIME% MSM Script has completed >> %LOGPATH%\%LOGFILE%
echo %CUR_DATE% %TIME% MSM Script has completed
echo ------------------------------------------ >> %LOGPATH%\%LOGFILE%
echo ------------------------------------------
color
cls