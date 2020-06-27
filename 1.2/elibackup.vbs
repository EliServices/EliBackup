PGN = "EliBackup"
SF = "Systemfehler. Bitte kontaktieren sie " & vbNewLine & "den Softwareherrausgeber!" & vbNewLine & vbNewLine
FE = PGN & " - Fehlermeldung"
ZE = Chr(34)
opt = 4 'Number of Options (general)
lenopt = 13 'Length of optiontext before value

ReDim GSett(opt) 'General Settings Integer 

Set objNet = CreateObject ("WScript.NetWork")
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject( "WScript.Shell" )
Set objShellApp  = CreateObject( "Shell.Application" )

Set objWMIService = GetObject("winmgmts:\\" & "." & "\root\cimv2") 
Set colItems = objWMIService.ExecQuery _  
("Select * From Win32_DisplayConfiguration")

defaultLocalDir = shell.ExpandEnvironmentStrings ("C:\")
Function ChooseFile (ByVal initialDir)
	Set ex = shell.Exec( "mshta.exe ""about: <input type=file id=X><script>X.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(X.value);close();resizeTo(0,0);</script>""" )
    ChooseFile = Replace( ex.StdOut.ReadAll, vbCrLf, "" )
    Set ex = Nothing
End Function

DefaultFolder = ""
Function SelectFolder( DefaultFolder )
    Set objFolder = objShellApp.BrowseForFolder( 0, "Backup-Ordner auswählen:", 0, DefaultFolder )
    SelectFolder = objFolder.Self.Path
End Function

Pathname = fso.GetParentFolderName(WScript.ScriptFullName)
config = Pathname & "\elibackup.config"

'Finished Setup ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ General setup done ---

x = 0
If (fso.FileExists(config)) Then
    Set f = fso.OpenTextFile(config, 1)
    Do While f.AtEndOfStream <> True 
        x = x + 1
        ReDim Preserve Settings(x) 
        org_zeile = f.Readline  
        Settings(x) = org_zeile
    Loop
    f.Close
Else 
	ERO = MsgBox (SF & "Fehlercode: 001",16,FE)
	WScript.Quit  
End If

'Read configuration file --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

sInt = 0 'Number of settings found
fInt = 0 'Number of files found
dInt = 0 'Number of directorys found
For i = 1 To x
	scan = RTrim(LTrim(Settings(i)))
	If Left(scan,1) <> "#" And Left(scan,1) <> "" Then	
		Select Case Left(scan,lenopt)
		Case "Debugging  = "
			GSett(1) = scan
			sInt = sInt + 1
		Case "Autostart  = "
			GSett(2) = scan
			sInt = sInt + 1
		Case "Choosedest = "
			GSett(3) = scan
			sInt = sInt + 1
		Case "Backuppath = "
			GSett(4) = scan
			sInt = sInt + 1
		Case Else 
			If Mid(scan,2,2) = ":\" Then 
				dsearch = InStrRev(scan,"\") + 1	
				fsearch = InStr(dsearch,scan,".")
				Select Case fsearch
				Case 0
					dInt = dInt + 1
					ReDim Preserve dpath(dInt)
					dpath(dInt) = scan		
				Case Else 
					fInt = fInt + 1
					ReDim Preserve fpath(fInt)
					fpath(fInt) = scan					
				End Select
			End if
		End Select		
	End If
Next

'Scanned config file ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

For i = 1 To opt -1
	l = Len(GSett(i)) - lenopt
	If Right(GSett(i),l) = "enabled" Then
	
		Select Case i
		Case 1
		Case 2
			'enable Autostart
		Case 3
			backuppath = SelectFolder( DefaultFolder ) 
			If Left(backuppath,1) <> "\" Then
				backuppath = backuppath & "\"
			End If
			
			If fso.FolderExists(backuppath) Then
			
			Else
				Ausg = MsgBox(SF & "Fehlercode: 201" & vbCrLf & vbCrLf & ZE & "Folder does not exist" & ZE,16,SE)
				WScript.Quit()
			End If
		End Select
		
	ElseIf Right(GSett(i),l) = "disabled" Then
	
		Select Case i
		Case 1
		Case 2
			'disable Autostart
		Case 3
			lb = Len(GSett(4)) - lenopt
			backuppath = Right(GSett(4),lb)		
			If backuppath = "" Then
				Ausg = MsgBox(SF & "Fehlercode: 202" & vbCrLf & vbCrLf & ZE & "No backuppath given." & ZE,16,SE)
				WScript.Quit()
			End If
		End Select
		
	Else
		Ausg = MsgBox(SF & "Fehlercode: 10" & CStr(i) & vbCrLf & vbCrLf & ZE & "Can not read " & CStr(i) & ". setting of .config file."& ZE,16,SE)
		WScript.Quit()
	End If 
Next

'Settings Done ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Conclusion = "Statusbericht:" & vbNewLine & vbNewLine & "Einzeldateien:"
For i = 1 To fInt
	If fso.FileExists(fpath(i)) Then
		fso.CopyFile fpath(i), backuppath, True
		Conclusion = Conclusion & vbNewLine & fpath(i) & " - copied"
	Else
		Conclusion = Conclusion & vbNewLine & fpath(i) & " - Not found"
	End if	
Next

'Copied available files ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Conclusion = Conclusion & vbNewLine & vbNewLine & "Ordner:"
For i = 1 To dInt
	If fso.FolderExists(dpath(i)) Then

		If Right(dpath(i),1) = "\" Then
			dpath(i) = Left(dpath(i),Len(dpath(i)) - 1)
		End If
		
		fso.CopyFolder dpath(i), backuppath, True
		Conclusion = Conclusion & vbNewLine & dpath(i) & " - copied"
	Else
		Conclusion = Conclusion & vbNewLine & dpath(i) & " - Not found"
	End if	
Next

'Copied available folders -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Endmessage = MsgBox(Conclusion,64,PGN)

'Finished backup ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ Backup done ---