@ECHO off
title Printer System Cleanup

::Stopping the Print Spooler
	net stop spooler
	ECHO.
	ECHO Cleaning up the registry a bit, Please Wait ...
	ECHO.

::Removing Network Printer Registry Keys
	ECHO Y | reg delete HKEY_CURRENT_USER\Printers\Connections\
	ECHO.
	ECHO Registry cleanup completed, starting print spooler.
	ECHO.

::Starting the Print Spooler
	net start spooler
	ECHO.
	ECHO Printers are on their way, this'll take a minute or two.
	ECHO.

::Forcing a GPUPDATE
	ECHO N | gpupdate /force
	ECHO.

::Opening up the devices and printers window, and loading the printers to confirm for the end user. 

	ECHO Printers are now added, please confirm by checking the recently opened window.
	control /name Microsoft.DevicesAndPrinters  
	ECHO.
	ECHO.
	ECHO.
	ECHO Make sure to set your DEFAULT printer!!
	ECHO.
	ECHO Done!
	ECHO.