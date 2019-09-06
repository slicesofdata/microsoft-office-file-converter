;------------------------------------------------------------------------------;
; Namespace:      thelastsliceofdata 
; AHK version:    AHK 1.1.30.00 
; Function:       Bulk Office file converter
; Language:       English
; Tested on:      Win 7 (U64)
; Version:        0.1.05.00
; Date:           08/29/2019

; thanks to https://github.com/ahkon and other for providing useful documention

; PowerPoint extension codes: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb251061(v=office.12)
; https://docs.microsoft.com/en-us/office/vba/api/PowerPoint.Presentation.SaveAs
;------------------------------------------------------------------------------;; 
;-------------------------------------------------------------------------------
;
; This script converts files between versions. MS Office needs to be installed to run.
;
;-------------------------------------------------------------------------------
#SingleInstance, Force
#NoEnv
ListLines Off

;IfExist, %A_scriptDir%\bin\MS FileConverter.ico
Menu, Tray, Icon, Shell32.dll, 135
;#Include X:\Progs\AutoHotKey\Lib\COM.ahk
;#Include X:\Progs\AutoHotKey\Lib\Acc.ahk
;#Include X:\Progs\AutoHotKey\Lib\Excel.AHK

;dirs := A_ProgramFiles . (A_PtrSize=8 ? " (x86)" : ""), A_ProgramFiles

;**************************
;* Set some parameters    *
;**************************
; set the directories to be searched for office programs
dirs  := A_ProgramFiles . (A_PtrSize=8 ? " (x86)" : "") . "\Microsoft Office," . "C:\Program Files\Microsoft Office" . "," . "C:\Program Files (x86)\Adobe" . "," . "C:\Program Files\Adobe" ; directory do search for programs.
;msgbox,,,% dirs


; an array of program details: conversion from ext, conversion to ext, and program exe names
;;;;;;; should be retested with fake program names
;ProgramTable := [["doc","docx","winword.exe"]
;                ,["xls","xlsx","excel.exe"]
;			 ,["ppt","pptx","powerpnt.exe"]
;			 ,["xxx","xxx","xxx.exe"]]
ProgramTable :=  [["doc","docx","winword.exe"]
			 ,["docx","doc", "winword.exe"]
                ,["doc","pdf", "acrobat.exe"]
			 ,["docx","pdf","acrobat.exe"]
			 ,["ppt","pptx","powerpnt.exe"]
			 ,["ppt","pdf","acrobat.exe"]
			 ,["pptx","pdf","acrobat.exe"]
			 ,["xls","xlsx","excel.exe"]]

; a display message
SplashTextOn,350,30,SillyScriptz MS Office File Converter,Checking available programs, please wait...

numPrograms := ProgramTable.Length() ;msgbox,,,%numPrograms%

; then check for existence of programs installed and get a list if they exist
loop %numPrograms% ; it loops by counting down
{
	prog_names := ProgramTable[A_Index,3] . "," . prog_names
}
plist := programs_exist(prog_names,dirs) ; returns only the existing programs
SplashTextOff
loop, parse, plist, `;
{
	if not instr(l,A_loopfield)
		l := A_loopfield . ";" . l
}	
plist := SubStr(plist, 1, strlen(plist)-1) ; remove the last character ;
plist := SubStr(l, 1, strlen(l)-1) ; remove the last character ;
SplashTextOff



;***************************
;*   parts for the GUI     *
;***************************
Loop, Parse, plist, `; ; then get program count for warning messages
{
	pcount := A_Index
}
If (pcount = "") | (pcount < 1) ; then present a message
{
	Msgbox,0x30,Warning,There are NO programs installed on this computer for converting files. Install the programs first and then run this program.`n`nThe following typically programs for converting the files are not installed.`n`n%plist%`n`n, ;30
} else {
	Msgbox,0x24,Message,There are %pcount% programs installed for converting files. Those programs are:`n`n%plist%.`n`nDo you wish to proceed with converting files associated with these programs? If not, install the programs associated with your conversion and try again. 
	IfMsgBox,No,ExitApp
}

; loop through the columns and assign extensions and programs to csv strings
loop %numPrograms% ; 
{
	prog_from_exts := ProgramTable[A_Index,1] . "," . prog_from_exts 
	prog_to_exts   := ProgramTable[A_Index,2] . "," . prog_to_exts
	prog_names     := ProgramTable[A_Index,3] . "," . prog_names
}	
; then clean them up by removing the last comma
prog_from_exts := SubStr(prog_from_exts, 1, strlen(prog_from_exts)-1) 
prog_to_exts   := SubStr(prog_to_exts, 1, strlen(prog_to_exts)-1) 
prog_names     := SubStr(prog_names, 1, strlen(prog_names)-1) 

; get row and column count of table of programs to reference extensions  ?????????? remove????????
for RowIndex, Row in ProgramTable
{
	rowcount := A_Index
	
	temp := (RowIndex = 1 ? "[[" : " [")
	for ColumnIndex, Column in Row
		
		;msgbox,,,column index = %columnindex% 
		colcount := A_index
		;MsgBox,,,%A_Index%
		temp .= Column ", "
		;msgbox,,,column = %column%
	output .= SubStr(temp, 1, strlen(temp)-2) "]`n"
}

;*******************
;*   Build GUI    *
;*******************
gui -MinimizeBox
gui Margin,6,6
gui Font,s10
gui Add,Text,xm y10 w300 h150,
gui,Add,Text,xs+8 yp+10,File Extensions to convert (check all that apply):

;gui Add,Checkbox,xp+10 yp+20 h20 Section Checked vOpt1,%ext1%   ; to have it checked by default
loop %numPrograms% ; make the gui dynamic relative to the programs in the plist
{
	if instr(plist, ProgramTable[A_Index,3]) ; only if they exist
	{
		option := "." . ProgramTable[A_Index,1] . " to ." . ProgramTable[A_Index,2] ;. " " . ProgramTable[A_Index,3] 
		gui Add,Checkbox,xs+10 yp+20 h20 vOpt%A_Index%, %option% ; add new checkbox below
	}
}	
gui Add,Button,xs+50 y+30 w100 h30 vConvertButton gConvert,Convert
gui Add,Button,x+10 wp hp vCloseButton gClose,Close
;gui Add,Button,x+10 wp hp gGUITest,GUI Test
gui Show,,SillyScript MS Office File Converter
return


;*******************
;*    GUI Close    *
;*******************
GUIEscape:
GUIClose:
ExitApp
return


;***************
;* converting  *
;***************
Convert: ; this label content executes if the convert button is pressed 
Gui, Submit, NoHide ;this command submits the guis' datas' state s
Gui, Hide

gosub, get_folder_and_files ; get the files in the selected folder

;Loop %ColCount%
Loop %numPrograms%
{
	If Opt%A_index% = 1 
	{
		from_ext  := ProgramTable[A_Index,1] ; index = row
		to_ext    := ProgramTable[A_Index,2] 
		prog_name := ProgramTable[A_Index,3]
		;msgbox,,,from = %from_ext%`nto = %to_ext%`ncurrent = %prog_name%
		
		filez := list_files(ProjectFolder, from_ext, to_ext) 
		cnt := conversion_count(filez, from_ext, to_ext)

		if (cnt > 0 ) 
		{
		MsgBox,4,Message,You checked the box to convert %from_ext% -> %to_ext% in:`n`n%ProjectFolder%`n`n%cnt% file(s) in this folder to be converted.`n`nNote: If a file with the same name and the new extension (.e.g., %to_ext%) already exists, this file will NOT be converted because doing so would overwrite the existing file.`n`nDo you wish to proceed with the file conversion?
			IfMsgBox, Yes
			{
				conversion_loop(filez, from_ext, to_ext,1)
			}
		} else {
		MsgBox,,Message,You checked the box to convert %from_ext% -> %to_ext% in:`n`n%ProjectFolder%`n`n%cnt% file(s) in this folder to be converted.
		}
	} else {
		;MsgBox,0x30,Message,You never made a selection. Program will close.
			;ExitApp
	}	
}
 ; get the folder to convert, recursively

Gui, Hide
SplashTextOn, 300, 30, Message,Done: program closing!
Sleep, 2000
SplashTextOff
ExitApp
return

;-- Enable/Disable buttons
;GUIControl Disable,ConvertButton
;GUICOntrol Enable ,CloseButton

;**************
;*    stop    *
;**************
Close:
ExitApp
return

MsgBox,,,count = %fileCount%
SplashTextOn,200,30,Conversion,%FileCount% Files Converted!
sleep,500
SplashTextOff


;------------------------------------------------------------------------------------------------------------;
get_folder_and_files:
FileSelectFolder, ProjectFolder,%xlFileDir%,4,Select the Folder containing the files you want to convert.
SplitPath,ProjectFolder,folderName
if (ProjectFolder = "") {
   MsgBox,,Project Folder Selection Error, You didn't select a project folder.
	exitapp ;~ gosub, get_folder
} 
else {
;MsgBox,1,Project Selection,You selected the folder named: %folderName% 
	;IfMsgBox, Cancel
		;exitapp
	;IfMsgBox, ok
		;return
}

programs_exist(progs,dirs) {
	;SplashTextOn,350,30,MS File Converter,Gathering program list...
	;sleep, 1000
	plist =
	loop, parse, progs, `,
	{
		currentProg := A_LoopField
		;msgbox,,, the prog = %currentProg%
		Loop, parse, dirs, `,
		{
			Loop %A_loopfield%\*.exe,0,1
			{
			;tooltip,Searching for programs to complete conversion. Please wait momentarily:`n%theprog%,500,500
			;sleep,50
				If (A_LoopFileName = currentProg)
				{
					;SplashTextOn,350,30,MS File Converter: Gathering program list...,%A_LoopFileName%
				;MsgBox,,,Found: %A_LoopFileName%
					plist := A_LoopFileName . ";" . plist
				}
			}
		;tooltip
		}
	;	SplashTextOff	
	}
	return plist
}


list_files(directory, from_extension, to_extension)  ; https://www.autohotkey.com/boards/viewtopic.php?t=47428
{
	SplashTextOn,350,30,MS File Converter,Gathering file list...
	sleep, 1000
	files = 
	Loop %directory%\*.%from_extension%,0,1
	{
		if A_LoopFileAttrib contains H,R,S 
			continue
		
		SplitPath, A_LoopFileFullPath, fname, fdir, fext, fname_no_ext, drive
		;MsgBox,,,%fname%
		SplashTextOn,350,30,MS File Converter,%fname%
		If (fext = from_extension) ; remove similar file extensions
		{
			
			files = %A_LoopFileFullPath%`n%files%
		}
	}
	SplashTextOff
	return files
}

conversion_count(filelist, from_extension, to_extension) {
	filecount := 0
	loop, parse, filelist,`n  ; loop the list
	{
		IfExist, %A_LoopField%
		{
			SplitPath, A_loopfield, fname, fdir, fext, fname_no_ext, drive
			 
			If (%fext% == %to_extension%)
			{
				FileCount++ ; := filecount + 1	
			}
		}
	}
	return FileCount ;A_index ;this will also count the new extension files. so this needs editing. 
}

conversion_loop(filelist, from_extension, to_extension, conversion) {
	
	loop, parse, filelist,`n  ; loop the list
	{
		IfExist,%A_LoopField%
			IfEqual,conversion,1 
				file_conversion(A_LoopField, from_extension, to_extension) ; perform conversion	
	}
}

file_conversion(full_path, from_extension, to_extension) {
	;msgbox,,,from = %from_extension%`nto = %to_extension%`nfullpath = %full_path%
	
	SplitPath, full_path, fname, fdir, fext, fname_no_ext, drive	;msgbox, %A_LoopFileFullPath%
	
	to_file := fdir . "\" . fname . to_extension	;Msgbox,,,to = %to_extension%
	IfNotExist,%to_file%	; only convert if a newer extension does not exist already
	{
		If (from_extension = "ppt") | (from_extension = "pptx") 
		{
			ppt2_convert(full_path, from_extension, to_extension)
		}	
		IfEqual,from_extension,xls
			xl_convert(full_path, from_extension, to_extension)
		
		If (from_extension = "doc") | (from_extension = "docx") 
		{
			word_convert(full_path, from_extension, to_extension)
		}
	}
}

save_message(directory,name_no_ext,fromext) {
	mousegetpos, x, y
	tooltip, Converting: %directory%\%name_no_ext%.%fromext%, (x + 20), (y + 20) ;, 1
	sleep, 500
}


xl_convert(filepath, from_extension, to_extension) {
	;msgbox, Filepath = %filepath%
	SplitPath, filepath, fname, fdir, fext, fname_no_ext, drive
	
	XL := ComObjCreate("Excel.Application") 	; Creates excel object
	XL.DisplayAlerts := False ; this is Set to False to suppress prompts and alert messages 
	XL.Visible := False ; set to false if you don't need excel to be seen
	
	Xl_Workbook := Xl.Workbooks.Open(filepath) ; open an existing file AND ALSO get a handle to the current workbook.
	;if (from_extension != "xlsx") 
	if (to_extension = "csv") {
		newfile := fdir . "\" . fname_no_ext . "." . to_extension
		IfNotExist, %newfile%
		{
			XL_Workbook.SaveAs(newfile, 6) ;6) write out the csv file for temporary mailing use
			save_message(fdir,fname_no_ext,to_extension) 
		}
	}
	if (to_extension = "xlsx") {
		newfile := fdir . "\" . fname_no_ext . "." . to_extension
		IfNotExist, %newfile%
		{
			XL_Workbook.SaveAs(newfile, 51) ;51=2007(xlsx)
			save_message(fdir,fname_no_ext,to_extension) 
		}
	}	
	if (to_extension = "xls") {
		newfile := fdir . "\" . fname_no_ext . "." . to_extension
		IfNotExist, %newfile%
		{
			XL_Workbook.SaveAs(newfile, 56) ;56=2003(xls)
			save_message(fdir,fname_no_ext,to_extension) 
		}
	}
	XL_Workbook.Close(1) 
	XL.Quit ; should include ()?
	XL := "" ; needed in ahk to fully close out of PowerPoint. 
	
	tooltip
}

ppt_convert(filepath, from_extension, to_extension) {
	
	SplitPath, filepath, fname, fdir, fext, fname_no_ext, drive
	;msgbox, FROM:`n%fdir%\%fname%`n`nTO:`n%fdir%\%fname_no_ext%.%to_extension%
	
	if (from_extension = "ppt") {
		PPT := ComObjCreate("PowerPoint.Application")
		;PPT.DisplayAlerts := False ; this is Set to False to suppress prompts and alert messages 
		;PPT.Visible := False ; set to false if you don't need excel to be seen
		
		PPT_Pres := PPT.Presentations.Open(filepath, false,false,false)
			;PPT_Pres := PPT_Pres.ActivePresentation
		newfile := fdir . "\" . fname_no_ext . "." . to_extension
		IfNotExist, %newfile%
		{
			PPT_Pres.SaveAs(newfile, 11)
			save_message(fdir,fname_no_ext,to_extension) 
		}
		
		PPT_Pres.Close ; should include ()?
		PPT.Quit ; should include ()?
		PPT := "" ; needed in ahk to fully close out of PowerPoint. 
		
		tooltip
	}
}

ppt2_convert(filepath, from_extension, to_extension) { ; https://www.autohotkey.com/boards/viewtopic.php?t=37956
	;msgbox, Filepath = %filepath%
	SplitPath, filepath, fname, fdir, fext, fname_no_ext, drive
	;msgbox, FROM:`n%fdir%\%fname%`n`nTO:`n%fdir%\%fname_no_ext%.%to_extension%
	
	if (from_extension = "ppt") | (from_extension = "pptx") {
		PPT := ComObjCreate("PowerPoint.Application")
		;PPT.DisplayAlerts := False ; this is Set to False to suppress prompts and alert messages 
		;PPT.Visible := False ; set to false if you don't need excel to be seen
		
		PPT_Pres := PPT.Presentations.Open(filepath, false, false, false)
				
		newfile := fdir . "\" . fname_no_ext . "." . to_extension
		IfNotExist, %newfile%
		{
			if (to_extension = "pptx") ; 
			{
				PPT_Pres.SaveAs(newfile, 11) ; to pptx
			}
			if (from_extension = "ppt") & (to_extension = "pdf")
			{
				PPT_Pres.SaveAs(newfile, ppSaveAsPDF := 32) ; coverts to pdf
			}
			if (from_extension = "pptx") & (to_extension = "pdf")
			{
				PPT_Pres.SaveAs(newfile, ppSaveAsPDF := 32) 
			}
			
			save_message(fdir,fname_no_ext,to_extension) 
		}
		
		PPT_Pres.Close ; should include ()?
		PPT.Quit ; should include ()?
		PPT := "" ; needed in ahk to fully close out of PowerPoint. 
		
		tooltip
	}
}


word_convert(filepath, from_extension, to_extension) {
;https://github.com/ahkon/MS-Office-COM-Basics/blob/master/Examples/Word/Bookmarks.ahk
	
	SplitPath, filepath, fname, fdir, fext, fname_no_ext, drive
	
	if (from_extension = "doc") | (from_extension = "docx") {
		;tooltip, saveWordFile, 900, 10
				
		WORD := ComObjCreate("Word.Application")
		;WORD.Visible := 1 ;0=invisible; visible= 1
		WORD_Doc := WORD.Documents.Open(FileName := filepath) ;, ReadOnly:=True
		
		newfile := fdir . "\" . fname_no_ext . "." . to_extension  ; Save the document.
		;MsgBox,,,%newfile%
		IfNotExist, %newfile% 
		{
			if (from_extension = "doc") & (to_extension = "docx")
			{
				WORD_Doc.SaveAs(newfile, wdFormatDocumentDefault := 16)  ; Save the document as docx.
			}
			if (from_extension = "doc") & (to_extension = "pdf")
			{
				WORD_Doc.SaveAs2(newfile, wdSaveAsPDF := 17) ; 17 = pdf?wdFormatPDF := 17
			}
			if (from_extension = "docx") & (to_extension = "doc")
			{
				WORD_Doc.SaveAs(newfile, wdFormatDocument97 := 0)  ; Save the document as old doc.
			}
			if (from_extension = "docx") & (to_extension = "pdf")
			{
				WORD_Doc.SaveAs2(newfile, wdSaveAsPDF := 17 ) ;wdSaveFormat := ) ; 17 = pdf?
			}
			save_message(fdir,fname_no_ext,from_extension)
		}
		WORD_Doc.Close()  ; Close the document.
		WORD.Quit()  ; Quit Word.
		WORD := "" 
		
		tooltip
	}
}


close_program(progname){

	Process, Exist, %progname%
	If ErrorLevel <> 0
		MsgBox,0x34,,%progname% is currently open and running. In order to convert files safely, please first close the program.`n`nDo you wish to close %progname% and continue with conversion?
		IfMsgBox,Yes
			Process, Close, %progname%
		Else
			{
				MsgBox,,Message,File Converter will close now.
				Exitapp
			}
}


msg(mess = "Test message", title = "Title", w = 200, h = 30, dur = 1000){
	SplashTextOn,%w%,%h%,%title%,%mess%
	sleep, %dur%
	SplashTextOff
}


; for examining the program table Accepts a 2 dimensional array. Anything else will probably print garbage..   
PrintTable(table)
{
	for RowIndex, Row in table
	{
		temp := (RowIndex = 1 ? "[[" : " [")
		for ColumnIndex, Column in Row
			temp .= Column ", "
		output .= SubStr(temp, 1, strlen(temp)-2) "]`n"
	}
}

RemoveToolTip:
ToolTip
return


;*********************** run delete files ********************************.
deletefile:
FileDelete, %fdir%\%fname%
Return
;*********************** run delete files ********************************.