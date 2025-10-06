# Copy Autoit

[AutoIt sketch](https://www.autoitscript.com/site/autoit/downloads/) that copies changed files from source to dest. 
When added to Taskscheduler it will make a backup of the changed files. 

<img width="450" alt="image" src="https://github.com/user-attachments/assets/4ff8c2f5-3cab-4ef2-97b2-227e42cd7e50" />

Start the exe-files with Taskscheduler.
The interupting mode starts a windoes in which the copied files are displayed. This is annoying when you are typing and the windows takes takes over.
With the "silent mode" sketch the backup process is running in the background.

In the CopyFilesToA.au , ToB et cetera the source and destination folders are hardcoded
The CopyFilesTo.au3 is generic and the source and destination folders are places in the arguments of the by autoIt compiled exe-file.

### CopyFilesToA.au3 --> CopyFilesToA.exe  Interupting mode
Simple start of the sketch with MSGBoxes interupting your work but handy to see that it is working
In the CopyFilesToA , ToB et cetera the source and destination folders are hardcoded 
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
### CopyFilesTo.au3 --> CopyFilesTo.exe Silent mode 
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

<img width="450" alt="image" src="https://github.com/user-attachments/assets/0cf45372-ac2c-4164-b1d0-d47de46b4531" />
Example: Interupting mode


<img width="450" alt="image" src="https://github.com/user-attachments/assets/ec05c965-256a-457c-bd2b-f68b56764765" />
Example: Silent mode


In the "add arguments" field type the: Source Dest and /min
c:\Users\XXXXXX\Documents\Files\ h:\ /min

c:\Users\XXXXXX\Documents\Files\ "B:\DataEd\My Documents\Files\" /min
/min suspresses the Messageboxes

### Always perform a "CHECK NAMES" in Taskscheduler before saving.
Otherwise you get an error.

<img width="450" alt="image" src="https://github.com/user-attachments/assets/95a9fc1f-058d-44b1-a198-b832d7c54077" />
