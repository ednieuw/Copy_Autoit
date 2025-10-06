#include <Array.au3>
#include <GuiEdit.au3>
#include <GuiConstantsEx.au3>

$MsgBoxStr = "EJN Aug 2025 Copy Files To"                                                              ; The title of the message box.
                                      ;;         MsgBox(0, $MsgBoxStr, "Started CopyFilesTo ",1)

$MaxToCopyFiles = 1000                                                                                           ;Let's not copy more than 1000 files

$Nodisplay= True

If $CmdLine[0] <3 Then
    MsgBox(0, $MsgBoxStr,"ERROR  No source and dest parameters",1)
    Sleep(5000)
    Exit 
EndIf
                                                                                                                                ; if /min is added the script will run minimized
If $CmdLine[0] > 2 And $CmdLine[3] = "/min" Then                                                     ; if /min added no interuptions from the messagebox
    WinSetState(@ScriptName, "", @SW_MINIMIZE)
	$Nodisplay=True 
EndIf

$Nodisplay= True

if $CmdLine[0] > 1 Then                                                                                           ; if two command line parameters available ie source and dest
	$CopyToDest = StringLeft($CmdLine[2] , 3)
		$MsgBoxStr = $MsgBoxStr  & $CmdLine[1] & " to" & $CmdLine[2]
		If $Nodisplay = False Then MsgBox(0, $MsgBoxStr, "Copy from" & $CmdLine[1] & " to " & $CmdLine[2] , 5)
	$MsgBoxStr = $MsgBoxStr   		
EndIf

if $CopyToDest = "" Or $CmdLine[1] ="" Then                                                           ; if destination drive is empty
  MsgBox(0, $MsgBoxStr, "ERROR No source and dest parameters", 5)
  Exit 
EndIf

$try=0
                                       ;;   MsgBox($MB_SYSTEMMODAL, $MsgBoxStr "Started waiting for A:",1)
Do
	Sleep(99)
    $try = $try + 1
    If $try >= 200 Then 
		MsgBox($MB_SYSTEMMODAL, $MsgBoxStr, $CopyToDest & "ERROR Backup drive not ready",5) 
		Exit
	EndIf	
Until FileExists($CopyToDest)
_CopyFolderNoOverwrite($CmdLine[1] , $CmdLine[2])                                                 ; kopieer en overschrijf oudere bestanden

                                      ;;MsgBox($MB_SYSTEMMODAL, $MsgBoxStr, "Copy Ended",4)

Func _CopyFolderNoOverwrite($SourceFolder, $DestinationFolder)
    Local $Array = _FileListToArray_Recursive($SourceFolder, "*", 1, 2, 1)
    If Not IsArray($Array) Or Not $Array[1] Then Return                                               ;MsgBox(4096, 'No Files Found', 'Cannot Find any Files in ' & $SourceFolder,1)
    Local $TCopyArray[$Array[0] + 1][2]
	Local $CopyArray[$Array[0] + 1][2]
    $TCopyArray[0][0] = $Array[0]                                                                            ; de geselecteerde files opslaan
    For $i = 1 To $Array[0]
        $TCopyArray[$i][0] = $Array[$i]
        $TCopyArray[$i][1] = StringRegExpReplace($Array[$i], StringReplace($SourceFolder, "\", "\\"), StringReplace($DestinationFolder, "\", "\\"))
    Next
	                                                                                                                      ; Alle bestanden hebben nu de juiste source en destination strings. Nu kunnen wij de tijden gaan vergelijken.
	local $j
	$j=1
	$CopyArray[0][0]=0
	    For $i = 1 To $Array[0]
		if FileGetTime($TCopyArray[$i][0],0,1) > FileGetTime($TCopyArray[$i][1],0,1) Then
        $CopyArray[$j][0] = $Array[$i]
        $CopyArray[$j][1] = StringRegExpReplace($Array[$i], StringReplace($SourceFolder, "\", "\\"), StringReplace($DestinationFolder, "\", "\\"))
		$CopyArray[0][0]=$j
		$j=$j+1
		endif
    Next
	If $Nodisplay = False And $CopyArray[0][0]=0 Then MsgBox(4096,$MsgBoxStr, "Cannot Find any Files to Copy in " & $SourceFolder, 1)
	If $CopyArray[0][0] = 0 Then Return 
		;
    If $CopyArray[0][0] > $MaxToCopyFiles then  $CopyArray[0][0] = $MaxToCopyFiles
                                  ;_ArrayDisplay($CopyArray)                                                          ; 2 Dimensional Array with Source File Path and Destination File Path
	Local $hEdit

	                                                                                                                                   ; Create GUI to print the copied files into
	If $Nodisplay = False Then 
		GUICreate($MsgBoxStr, 800, 300)
	    $hEdit = GUICtrlCreateEdit("Copied files" & @CRLF, 2, 2, 794, 268)
	   _GUICtrlEdit_SetLimitText($hEdit, 300000)                                                                 ;' 300000 chars in messagebox instead of default 30000. $hEdit,, -1) is unlimited size of the message box
       GUISetState()
    EndIf
    For $i = 1 To $CopyArray[0][0]                                                                                     ; Let's not copy more than 1000 files See line 5
		                     ;;MsgBox($MB_SYSTEMMODAL, 'File copied', 'File: ' & $CopyArray[$i][0],2)
		$result = FileCopy($CopyArray[$i][0], $CopyArray[$i][1], 9)
	    If $result = 0 Then                                                                                                   ; Handle specific read-only case
            FileSetAttrib($CopyArray[$i][1], "-R")                                                                    ; Remove read-only
            $result = FileCopy($CopyArray[$i][0], $CopyArray[$i][1], 9)
       EndIf	
	If $Nodisplay = False Then 	_GUICtrlEdit_AppendText($hEdit, $i & ":" & $CopyArray[$i][0] & @CRLF)
    Next

	If $Nodisplay = False Then 
		_GUICtrlEdit_AppendText($hEdit, @CRLF & "Last file copied. Will close in 5 seconds"& @CRLF)
	   Sleep(5000)
       GUIDelete()
    EndIf
EndFunc   

;===============================================================================
; $iRetItemType: 0 = Files and folders, 1 = Files only, 2 = Folders only
; $iRetPathType: 0 = Filename only, 1 = Path relative to $sPath, 2 = Full path/filename
Func _FileListToArray_Recursive($sPath, $sFilter = "*", $iRetItemType = 0, $iRetPathType = 0, $bRecursive = False)
    Local $sRet = "", $sRetPath = ""
    $sPath = StringRegExpReplace($sPath, "[\\/]+\z", "")
    If Not FileExists($sPath) Then Return SetError(1, 1, "")
    If StringRegExp($sFilter, "[\\/ :> <\|]|(?s)\A\s*\z") Then Return SetError(2, 2, "")
    $sPath &= "\|"
    $sOrigPathLen = StringLen($sPath) - 1
    While $sPath
        $sCurrPathLen = StringInStr($sPath, "|") - 1
        $sCurrPath = StringLeft($sPath, $sCurrPathLen)
        $Search = FileFindFirstFile($sCurrPath & $sFilter)
        If @error Then
            $sPath = StringTrimLeft($sPath, $sCurrPathLen + 1)
            ContinueLoop
        EndIf
        Switch $iRetPathType
            Case 1 ; relative path
                $sRetPath = StringTrimLeft($sCurrPath, $sOrigPathLen)
            Case 2 ; full path
                $sRetPath = $sCurrPath
        EndSwitch
        While 1
            $File = FileFindNextFile($Search)
            If @error Then ExitLoop
            If ($iRetItemType + @extended = 2) Then ContinueLoop
            $sRet &= $sRetPath & $File & "|"
        WEnd
        FileClose($Search)
        If $bRecursive Then
            $hSearch = FileFindFirstFile($sCurrPath & "*")
            While 1
                $File = FileFindNextFile($hSearch)
                If @error Then ExitLoop
                If @extended Then $sPath &= $sCurrPath & $File & "\|"
            WEnd
            FileClose($hSearch)
        EndIf
        $sPath = StringTrimLeft($sPath, $sCurrPathLen + 1)
    WEnd
    If Not $sRet Then Return SetError(4, 4, "")
    Return StringSplit(StringTrimRight($sRet, 1), "|")
EndFunc   ;==>_FileListToArray_Recursive