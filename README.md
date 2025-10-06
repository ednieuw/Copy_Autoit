# Copy_Autoit
AutoIt sketch that copies changed files from source to dest .
Start the exe files with taskscheduler. 

### interupting mode
Simple start of the sketch with MSGBoxes interupting your work but Handy to see that it is working
```
MsgBox(0, $MsgBoxStr "Started waiting for A:",1)
Do
	Sleep(99)
    $try = $try + 1
    If $try >= 200 Then 
		MsgBox($MB_TASKMODAL, $MsgBoxStr, "Backup drive A: not ready",3) 
		 Exit
	EndIf	
Until FileExists("A:\")
_CopyFolderNoOverwrite("c:\Users\ednie\Documents\Files\" , "A:\DataEd\My Documents\Files\")    ; kopieer en overschrijf oudere bestanden

MsgBox(0, $MsgBoxStr, "Copy Ended",4)

```
### Silent mode 
Runs in the background
```
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
```
