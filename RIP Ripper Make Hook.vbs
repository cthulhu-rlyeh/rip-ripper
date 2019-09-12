'''''''''''''''''''''''''''''''''''''''''''''''''''''
'   The RIP Ripper was created, from scratch, by
'   SrA Sibastian Bythewood during his time in
'   82 FSS/FSMPJ at Sheppard AFB.
'
'   Email: sibastian.bythewood@us.af.mil
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'	RIP Ripper Make Hook
'	-- This file created your customized hook.
'	Release 1.0.3
'	on 06 Apr 2015
'''''''''''''''''''''''''''''''''''''''''''''''''''''

set ofso = createobject("scripting.filesystemobject")

if not ofso.fileexists("RIP Ripper Core.vbs") then
	msgbox "Core not found. Please move this file into the same folder as the core to make your customized hook file. Once created, you can then move the newly created customized hook to any folder you need to rip RIPs in."
	wscript.quit
end if

set ofile = ofso.createtextfile("RIP Ripper Hook.bat",true)

ofile.write "@echo off" & vbcrlf
ofile.write "cls" & vbcrlf
ofile.write "if ""%~n0"" == ""RIP Ripper Hook"" (" & vbcrlf
ofile.write "	echo." & vbcrlf
ofile.write "	echo   DO NOT TOUCH ANYTHING UNTIL THIS WINDOW CLOSES" & vbcrlf
ofile.write "	echo       [except to enable macros or something]" & vbcrlf
ofile.write "	echo." & vbcrlf
ofile.write "	echo             YOU HAVE BEEN WARNED" & vbcrlf
ofile.write "	echo." & vbcrlf
ofile.write "	call """ & ofso.getabsolutepathname(".") & "\RIP Ripper Core.vbs"" ""%~dp0""" & vbcrlf
ofile.write ") else (" & vbcrlf
ofile.write "	echo." & vbcrlf
ofile.write "	echo Detected Ripper Hook file name change. Ripper stopped." & vbcrlf
ofile.write "	echo You should NOT be editing this file in any way..." & vbcrlf
ofile.write "	echo." & vbcrlf
ofile.write "	timeout /t 5 >nul 2>&1" & vbcrlf
ofile.write ")"
ofile.close

msgbox "Rip Ripper Hook created. Move/copy the hook into any folder with .BKP files and run it to have it hook the core and seperate/sort your RIPs." & vbcrlf & vbcrlf & "If someone has trouble running the hook, try having them rerun the hook maker from their computer to make a hook file personalized for them."