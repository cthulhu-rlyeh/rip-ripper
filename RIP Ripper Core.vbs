'''''''''''''''''''''''''''''''''''''''''''''''''''''
'   The RIP Ripper was created, from scratch, by
'   Sibastian Bythewood during his time in
'   82 FSS at Sheppard AFB.
'
'   Email: sibastian.bythewood.1@us.af.mil
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'	RIP Ripper Core
'	-- This file is the meat and potatoes of the ripper.
'	-- Do NOT modify the contents of this script. Please use the configuration files instead.
'	Release 2.1.2
'	on 06 Feb 2018
'''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' BEGIN INIT
'''''''''''''''''''''''''''''''''''''''''''''''''''''

const wdCell = 12
const wdCharacter = 1
const wdCharacterFormatting = 13
const wdColumn = 9
const wdItem = 16
const wdLine = 5
const wdParagraph = 4
const wdParagraphFormatting = 14
const wdRow = 10
const wdScreen = 7
const wdSection = 8
const wdSentence = 3
const wdStory = 6
const wdTable = 15
const wdWindow = 11
const wdWord = 2
const wdOrientLandscape = 1
const wdReplaceAll = 2
const wdPageView = 3

const ConReq = "CONCURRENT REQUEST"

public odoc, stype, sunit, sname

set oshell = createobject("wscript.shell")
set ofso = createobject("scripting.filesystemobject")
set oword = createobject("word.application")
oword.visible = false ' hide the doings

if wscript.arguments.count = 0 then
	msgbox "The RIP Ripper must be executed via the code hook method."
	wscript.quit
end if

scorefolder = ofso.getparentfoldername(ofso.getfile(wscript.scriptfullname))
sconfig = scorefolder & "\RIP Ripper Config.ini"

if len(freadini(sconfig,"general","types_config")) then
	stypesconfig = freadini(sconfig,"general","types_config") & "\RIP Ripper Types.ini"
else
	stypesconfig = scorefolder & "\RIP Ripper Types.ini"
end if

if len(freadini(sconfig,"general","raw_folder")) then
	sfolder = freadini(sconfig,"general","raw_folder") & "\"
else
	sfolder = wscript.arguments(0) & "\"
end if

set ofolder = ofso.getfolder(sfolder)
sbackups = sfolder & "Backups\"

if not ofso.fileexists(sconfig) then
	msgbox "Config file not found. Please verify that the file '~ RIP Ripper Config.ini' exists and is in the same folder as the Core file."
	wscript.quit
end if

if ofolder.files.count > freadini(sconfig,"general","file_count_allowed") and freadini(sconfig,"general","file_count_allowed") > 0 then
	msgbox "The number of files in the directory exceeds the established configuration threshold of " & freadini(sconfig,"general","file_count_allowed") & ". Please clean out the directory and remove old/unneeded RIPs before executing the RIP Ripper Hook again." & vbnewline & vbnewline & "The RIP Ripper will now terminate."
	wscript.quit
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' END INIT
'''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' BEGIN CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''

config_rip_ext = freadini(sconfig,"general","rip_ext")
config_skip_types = split(freadini(sconfig,"general","skip_types"),"|")
config_save_folder = freadini(sconfig,"general","save_folder")
config_use_unit_sort = freadini(sconfig,"general","use_unit_sort")
config_sort_folder = freadini(sconfig,"general","sort_folder")
config_use_type_sort = freadini(sconfig,"general","use_type_sort")
config_forced_types = split(freadini(stypesconfig,"forced","forced_types"),"|")

for each ofile in ofolder.files
	if lcase(ofso.getextensionname(ofile.path)) = config_rip_ext then
		if not ofile.attributes and 2 then
			sfile = ofile.name
			fcore ' This reaches to the main portion of the code, the "core"
		end if
	end if
next

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' END CORE
'''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' BEGIN OPTIONALS
'''''''''''''''''''''''''''''''''''''''''''''''''''''

if freadini(sconfig,"general","clear_tmp_files") then
	for each ofile in ofolder.files
		if lcase(ofso.getextensionname(ofile.path)) = "tmp" then
			if not ofile.attributes and 2 then
				ofso.deletefile(ofile.path)
			end if
		end if
	next
end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' END OPTIONALS
'''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' BEGIN RESOLUTION
'''''''''''''''''''''''''''''''''''''''''''''''''''''

oword.quit false

set ofso = nothing
set ofolder = nothing
set ofile = nothing
set oshell = nothing
set oword = nothing
set otar = nothing
set odoc = nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' END RESOLUTION
'''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' BEGIN FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''

' this function takes action once the raw file is found
' simply organizes the pieces in play and calls them in order
function fcore
	set otar = oword.documents.add(sfolder & sfile, , true)
	otar.activate
	oword.activewindow.view.type = wdpageview
	oword.activewindow.view.readinglayout = false

	bnorip = false

	fpagesetup
	fremoveart
	stype = ""
	fpulltype
	
	if stype = "RAR2NT" then
		oword.selection.pagesetup.orientation = wdOrientLandscape
	end if

	if bnorip = false then
		for each st in config_skip_types
			if st = stype then
				bnorip = true
			end if
		next
	end if

	do until bnorip = true
		otar.activate
		oword.selection.wholestory
		if len(oword.selection.text) < 200 then exit do
		
		fseprip
		fpullunit
		fpullname
		fsaverip
		if err.number <> 0 then ferrcall
	loop
	
	otar.close false
	fbackup
end function

' this function saves the end product with the proper naming convention
function fsaverip
	if len(sname) > 100 then sname = "ERROR_SNAME_LEN"
	if len(sunit) > 100 then sunit = "ERROR_SUNIT_LEN"
	if len(stype) > 20 then stype = "ERROR_STYPE_LEN"

	stype = fsaveable(stype)
	sunit = fsaveable(sunit)
	sname = fsaveable(sname)

	snewfile = stype & " - " & sunit & " - " & sname

	if len(config_save_folder) > 0 then
		xsavefolder = config_save_folder & "\"
	else
		xsavefolder = sfolder
	end if

	config_unit_sort = freadini(sconfig,"unit_sort",sunit)
	config_type_sort = split(freadini(sconfig,"type_sort",stype), "|")

	' this allows multiple save folders to be defined, separated by the | symbol
	for each ssavefolder in split(xsavefolder, "|")
		' this utilizes the specified sorting folder if unit sorting is enabled and available
		' only one unit sort folder allowed
		if config_use_unit_sort then
			if len(config_sort_folder) > 0 and len(config_unit_sort) > 0 then
				ssavefolder = config_sort_folder & "\"
				if not ofso.folderexists(ssavefolder) then ofso.createfolder ssavefolder
			end if
			if len(config_unit_sort) > 0 then
				ssavefolder = ssavefolder & "\" & config_unit_sort & "\"
				if not ofso.folderexists(ssavefolder) then ofso.createfolder ssavefolder
			end if
		end if

		' this sorts by type if enabled and available, regardless of unit sorting
		' allows three separate folders for types sorting, separated by | symbols
		stypefolder1 = ""
		stypefolder2 = ""
		stypefolder3 = ""
		if config_use_type_sort then
			for each stypefolder in config_type_sort
				if stypefolder1 = "" then 
					stypefolder1 = ssavefolder & stypefolder & "\"
					if not ofso.folderexists(stypefolder1) then ofso.createfolder stypefolder1
				elseif stypefolder2 = "" then 
					stypefolder2 = ssavefolder & stypefolder & "\"
					if not ofso.folderexists(stypefolder2) then ofso.createfolder stypefolder2
				elseif stypefolder3 = "" then 
					stypefolder3 = ssavefolder & stypefolder & "\"
					if not ofso.folderexists(stypefolder3) then ofso.createfolder stypefolder3
				end if
			next
		end if

		snewfileoriginal = snewfile

		' save after optionals
		if stypefolder1 <> "" then
			do until not ofso.fileexists(stypefolder1 & snewfile & ".docx")
				icnt = icnt + 1
				snewfile = snewfileoriginal & " " & icnt
			loop
			odoc.saveas stypefolder1 & snewfile & ".docx"
		end if
		if stypefolder2 <> "" then
			do until not ofso.fileexists(stypefolder2 & snewfile & ".docx")
				icnt = icnt + 1
				snewfile = snewfileoriginal & " " & icnt
			loop
			odoc.saveas stypefolder2 & snewfile & ".docx"
		end if
		if stypefolder3 <> "" then
			do until not ofso.fileexists(stypefolder3 & snewfile & ".docx")
				icnt = icnt + 1
				snewfile = snewfileoriginal & " " & icnt
			loop
			odoc.saveas stypefolder3 & snewfile & ".docx"
		end if

		if stypefolder1 = "" then 
			do until not ofso.fileexists(ssavefolder & snewfile & ".docx")
				icnt = icnt + 1
				snewfile = snewfileoriginal & " " & icnt
			loop
			odoc.saveas ssavefolder & snewfile & ".docx"
		end if
	next

	odoc.close false
end function

' this function cuts out everything from a string after the specified string
' returns either what's left or the whole string if not found
function ftrimafter(sstring, sfind)
	ftrimafter = sstring
	if instr(ftrimafter, sfind) > 0 then

	end if
end function

' this function is simply for error handling
' in a perfect world, this will never be needed
function ferrcall
	msgbox "Error: " & err.number & chr(13) _
		& "Source: " & err.source & chr(13) _
		& "Description: " & err.description _
		& chr(13) & chr(13) _
		& "Please report this error as soon as possible." _
		& chr(13) & chr(13) _
		& "sibastian.bythewood@us.af.mil"
	err.clear
end function

' this function simply creates a backup folder and saves the original raw rip into it
' this saves us from losing anything, even if it is ripped improperly
function fbackup
	if not ofso.folderexists(sbackups) then ofso.createfolder(sbackups)
	if ofso.fileexists(sbackups & sfile) then ofso.deletefile sbackups & sfile, true
	ofso.movefile sfolder & sfile, sbackups & sfile
end function

' this function removes the artifacts at the top of all rips created by EOM output
function fremoveart
	oword.selection.homekey wdstory
	if oword.selection.find.execute("PRIVACY ACT OF 1974") then
		oword.selection.moveup ,2
		oword.selection.homekey
		oword.selection.homekey wdstory,1
		oword.selection.delete
	end if
end function

' this function is supposed to properly format the rips
' it's small, but it often has issues when other things go wonky
function fpagesetup
	oword.activewindow.view.type = wdpageview
	oword.activewindow.view.readinglayout = false

	with oword.selection.pagesetup
		.TopMargin = oword.InchesToPoints(0.5)
		.LeftMargin = oword.InchesToPoints(0.5)
		.RightMargin = oword.InchesToPoints(0.5)
		.BottomMargin = oword.InchesToPoints(0.5)
		.HeaderDistance = oword.InchesToPoints(0.5)
		.FooterDistance = oword.InchesToPoints(0.5)
	end with
	
	oword.selection.wholestory
	oword.selection.font.size = "9"
	oword.selection.font.name = "Courier New"
end function

' this function converts the given string into a properly saveable string
' it removes things that windows doesn't like to be in file names
' and also helps a little with formatting of the file names
function fsaveable(sfilename)
	dim icnt
	fsaveable = sfilename
	icnt = 1

	fsaveable = trim(fsaveable)

	fsaveable = replace(fsaveable, "/", " ")
	fsaveable = replace(fsaveable, ":", " ")
	fsaveable = replace(fsaveable, "<", " ")
	fsaveable = replace(fsaveable, ">", " ")
	fsaveable = replace(fsaveable, "|", " ")
	fsaveable = replace(fsaveable, "\", " ")
	fsaveable = replace(fsaveable, "*", " ")
	fsaveable = replace(fsaveable, "?", " ")
	fsaveable = replace(fsaveable, vbnewline, " ")
	fsaveable = replace(fsaveable, vbcr, " ")
	fsaveable = replace(fsaveable, vbln, " ")
	fsaveable = replace(fsaveable, vbcrln, " ")
	
	fsaveable = replace(fsaveable, "    ", " ")
	fsaveable = replace(fsaveable, "   ", " ")
	fsaveable = replace(fsaveable, "  ", " ")
	fsaveable = replace(fsaveable, "  ", " ")

	sfilename = fsaveable
end function

' this function is supposed to find out what type of rip we're dealing with
' this is the function that is in the most danger if anything is changed on the rips
function fpulltype
	oword.selection.homekey wdstory
	With oword.selection.find
    		.clearformatting
    		.text = "Concurrent Request"
    		.replacement.clearformatting
    		.replacement.text = conreq
		.forward = true
		.execute ,,,,,,,,,,2
	End With
	
	oword.selection.homekey wdstory
	if oword.selection.find.execute(conreq) = false then bnorip = true
	
	if bnorip = true then
		for each rt in config_forced_types
			if oword.selection.find.execute(freadini(stypesconfig,"forced",rt)) = true then
				stype = rt
				bnorip = false
			end if
		next
	else
		oword.selection.homekey
		oword.selection.moveright wdword,1
		oword.selection.moveright wdword,1,1
		stype = trim(oword.selection.text)
	end if
end function

' this function seperates the rips and their pages based on the config
function fseprip
	ssepby = freadini(stypesconfig,"sepby",stype)
	
	if ssepby <> "EOF" then
		oword.selection.homekey wdstory
		oword.selection.find.execute ssepby
		oword.selection.moveright
		oword.selection.find.execute conreq
	else
		do until oword.selection.find.execute(conreq) = false
			oword.selection.moveright
		loop
	end if
	
	oword.selection.endkey
	oword.selection.delete ,2
	oword.selection.homekey wdstory,1
	
	oword.selection.cut
	set odoc = oword.documents.add
	odoc.activate
	oword.activewindow.view.type = wdpageview
	oword.activewindow.view.readinglayout = false

	oword.selection.paste
end function

' this function extracts the unit name per the config
function fpullunit
	sunitby = freadini(stypesconfig,"unitby",stype)
	
	if sunitby = "" then
		sunit = "ERROR_FPULLUNIT_CONFIG"
		exit function
	end if
	if left(sunitby, 3) = "***" then
		sunit = replace(sunitby, "***", "")
		exit function
	end if
	
	sfinder = ""
	imoveright = -1
	imoveleft = -1
	imoveup = -1
	imovedown = -1
	bpresshome = -1
	
	for each x in split(sunitby,"|")
		if sfinder = "" then sfinder = x else _
		if imoveright = -1 then imoveright = x else _
		if imoveleft = -1 then imoveleft = x else _
		if imoveup = -1 then imoveup = x else _
		if imovedown = -1 then imovedown = x else _
		if bpresshome = -1 then bpresshome = x
	next
	
	oword.selection.homekey wdstory
	oword.selection.find.execute sfinder
	if imoveright > 0 then oword.selection.moveright ,imoveright
	if imoveleft > 0 then oword.selection.moveleft ,imoveleft
	if imoveup > 0 then oword.selection.moveup ,imoveup
	if imovedown > 0 then oword.selection.movedown ,imovedown
	if bpresshome > 0 then oword.selection.homekey
	
	oword.selection.endkey ,1
	sunit = trim(oword.selection.text)
	
	sunit = split(sunit & "/", "/")(0)
	sunit = split(sunit & "(", "(")(0)
	sunit = split(sunit & ",", ",")(0)

	sunit = replace(sunit, " GROUP", " GP")
	sunit = replace(sunit, " WING", " WG")

	dim aunitend(7)
	aunitend(0) = " SQ"
	aunitend(1) = " GP"
	aunitend(2) = " WG"
	aunitend(3) = " RG"
	aunitend(4) = " RS"
	aunitend(5) = " EL"
	aunitend(6) = " DO"
	aunitend(7) = " CTR"

	for x = 0 to ubound(aunitend)
		if instr(sunit, aunitend(x)) > 0 then
			sunit = split(sunit, aunitend(x))(0) & aunitend(x)
		end if
	next
end function

' this function extracts the member's name per the config
function fpullname
	snameby = freadini(stypesconfig,"nameby",stype)
	
	if snameby = "" then
		sname = "ERROR_FPULLNAME_CONFIG"
		exit function
	end if
	if left(snameby, 3) = "***" then
		sname = replace(snameby, "***", "")
		exit function
	end if
	
	sfinder = ""
	imoveright = -1
	imoveleft = -1
	imoveup = -1
	imovedown = -1
	bpresshome = -1
	
	for each x in split(snameby,"|")
		if sfinder = "" then sfinder = x else _
		if imoveright = -1 then imoveright = x else _
		if imoveleft = -1 then imoveleft = x else _
		if imoveup = -1 then imoveup = x else _
		if imovedown = -1 then imovedown = x else _
		if bpresshome = -1 then bpresshome = x
	next
	
	oword.selection.homekey wdstory
	oword.selection.find.execute sfinder
	if imoveright > 0 then oword.selection.moveright ,imoveright
	if imoveleft > 0 then oword.selection.moveleft ,imoveleft
	if imoveup > 0 then oword.selection.moveup ,imoveup
	if imovedown > 0 then oword.selection.movedown ,imovedown
	if bpresshome > 0 then oword.selection.homekey
	
	oword.selection.moveright wdword,5,1
	sname = trim(oword.selection.text)
end function

' this is a weird function that reads the .ini config file
' i grabbed the idea for it from stackoverflow.com and heavily modified it
' it works, and quickly too, but it worries me
function freadini(sinifile, sarea, svar)
	set oinifile = ofso.opentextfile(sinifile,1,false)
	do while not oinifile.atendofstream
		sline = trim(oinifile.readline)
		if lcase(sline) = "[" & lcase(sarea) & "]" then
			sline = trim(oinifile.readline)
			do while left(sline,1) <> "["
				iequalpos = instr(sline,"=")
				if iequalpos > 0 then
					sleft = trim(left(sline,iequalpos-1))
					if lcase(sleft)=lcase(svar) then
						freadini = trim(mid(sline,iequalpos+1))
						exit do
					end if
				end if
				if oinifile.atendofstream then 
					freadini = ""
					exit do
				end if
				sline = trim(oinifile.readline)
			loop
		exit do
		end if
	loop
	oinifile.close
	set oinifile = nothing
end function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' END FUNCTIONS
'''''''''''''''''''''''''''''''''''''''''''''''''''''

' With Compliments of Sibastian Bythewood '