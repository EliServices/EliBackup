PGN = "EliBackup"
SF = "Systemfehler. Bitte kontaktieren sie " & vbNewLine & "den Softwareherrausgeber!" & vbNewLine & vbNewLine
FE = PGN & " - Fehlermeldung"
ZE = Chr(34)
opt = 4 'Number of Options (general)

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
    count = 0 
    Do While f.AtEndOfStream <> True 
        x = x + 1
        ReDim Preserve Settings(x) 
        org_zeile = f.Readline  
        Settings(x) = org_zeile
        count = count + 1
    Loop
    f.Close
Else 
	ERO = MsgBox (SF & "Fehlercode: 000",16,FE)
	WScript.Quit  
End If

'Read configuration file --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

l = Len(Settings(9)) - 12
debugging = Right(Settings(9),l) 'line 9, "Debugging = " => 12
If debugging = "enabled" Then
	deb = True
	Db = MsgBox("Debugging enabled",64,PGN & " - Debugger")
ElseIf debugging = "disabled" Then
	deb = false
Else
	Ausg = MsgBox(SF & "Fehlercode: 001" & vbCrLf & vbCrLf & ZE & "Can not read configuration file" & ZE,16,SE)
	WScript.Quit()
End If

'Setted up debug mode -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

l = Len(Settings(10)) - 12
autostart = Right(Settings(10),l) 'line 10, "Autostart = " => 12
If autostart = "enabled" Then
	'enable autostart
												If deb = True Then Db = MsgBox("Autostart enabled",64,PGN & " - Debugger") End if
ElseIf autostart = "disabled" Then
	'disable autostart
												If deb = True Then Db = MsgBox("Autostart disabled",64,PGN & " - Debugger") End if
Else
	Ausg = MsgBox(SF & "Fehlercode: 002" & vbCrLf & vbCrLf & ZE & "Can not read configuration file" & ZE,16,SE)
	WScript.Quit()
End If

'Setted up autostart option ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

l = Len(Settings(11)) - 13
choosedest = Right(Settings(11),l) 'line 11, "Choosedest = " => 13
If choosedest = "enabled" Then

												If deb = True Then Db = MsgBox("Choosemode enabled",64,PGN & " - Debugger") End If
	backuppath = SelectFolder( DefaultFolder ) 
	If Left(backuppath,1) <> "\" Then
		backuppath = backuppath & "\"
	End If
	If fso.FolderExists(backuppath) Then
												If deb = True Then Db = MsgBox("Der Ordner " & ZE & backuppath & ZE & " existiert.",64,PGN & " - Debugger") End If
	Else
		Ausg = MsgBox(SF & "Fehlercode: 004" & vbCrLf & vbCrLf & ZE & "Folder does not exist" & ZE,16,SE)
		WScript.Quit()
	End If
	
ElseIf choosedest = "disabled" Then
												If deb = True Then Db = MsgBox("Choosemode disabled",64,PGN & " - Debugger") End If
	
	l = Len(Settings(12)) - 13
	backuppath = Right(Settings(12),l) 'line 12, "Backuppath = " => 13
												If deb = True Then Db = MsgBox("Backuppath = " & backuppath,64,PGN & " - Debugger") End if
	If backuppath = "" Then
		Ausg = MsgBox(SF & "Fehlercode: 005" & vbCrLf & vbCrLf & ZE & "Can not read configuration file" & ZE,16,SE)
		WScript.Quit()
	End If
	
Else
	Ausg = MsgBox(SF & "Fehlercode: 003" & vbCrLf & vbCrLf & ZE & "Can not read configuration file" & ZE,16,SE)
	WScript.Quit()
End If

'Setted up backupfolder / Settings setted---------------------------------------------------------------------------------------------------------------------------------------------------------------------- Settings setup done ---

ReDim Preserve fpath(2)
fpath(1) = Settings(14 + opt) 'line 14 + opt
x = 15 + opt
i = 2
While Settings(x) <> ""
	ReDim Preserve fpath(i + 1)
	fpath(i) = Settings(x)
	i = i + 1
	x = x + 1
Wend
fcount = i - 1

If deb = True Then 
	For k = 1 To i
		Ausg = Ausg & vbNewLine & fpath(k)
	Next
	Db = MsgBox("fpath = " & Ausg,64,PGN & " - Debugger") 
End If

'Read filepaths -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

x = 20 + i + opt
ReDim Preserve dpath(2)
dpath(1) = Settings(x) 'from line 24 + i are the directorys
x = x + 1
ii = 2
While Left(Settings(x),2) <> " #"
	ReDim Preserve dpath(x + 1)
	dpath(ii) = Settings(x)
	ii = ii + 1
	x = x + 1
Wend
dcount = ii - 1

If deb = True Then 
	For k = 1 To ii
		Ausg2 = Ausg2 & vbNewLine & dpath(k)
	Next
	Db = MsgBox("dpath = " & Ausg2,64,PGN & " - Debugger") 
End If

'Read folderpaths ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

i = []
ii = []
x = []
xx = []
k = []
Ausg = []
Ausg2 = []

'Cleaned variables ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- Paths read ---

Conclusion = "Statusbericht:" & vbNewLine & vbNewLine & "Einzeldateien:"
For i = 1 To fcount
	If fso.FileExists(fpath(i)) Then
		fso.CopyFile fpath(i), backuppath, True
		Conclusion = Conclusion & vbNewLine & fpath(i) & " - copied"
												If deb = True Then Db = MsgBox("Dateipfad Nummer " & CStr(i) & " : " & vbNewLine & fpath(i) & " wurde nach " & backuppath & " kopiert.",64,PGN & " - Debugger") End if
	Else
		Conclusion = Conclusion & vbNewLine & fpath(i) & " - Not found"
	End if	
Next

'Copied available files ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Conclusion = Conclusion & vbNewLine & vbNewLine & "Ordner:"
For i = 1 To dcount
	If fso.FolderExists(dpath(i)) Then

		If Right(dpath(i),1) = "\" Then
			dpath(i) = Left(dpath(i),Len(dpath(i)) - 1)
												If deb = True Then Db = MsgBox("Ordnerpfad korrigiert (" & dpath(i) & ").",64,PGN & " - Debugger") End if 
		End If
		
		fso.CopyFolder dpath(i), backuppath, True
		Conclusion = Conclusion & vbNewLine & dpath(i) & " - copied"
												If deb = True Then Db = MsgBox("Ordnerpfad Nummer " & CStr(i) & " : " & vbNewLine & dpath(i) & " wurde nach " & backuppath & " kopiert.",64,PGN & " - Debugger") End if
	Else
		Conclusion = Conclusion & vbNewLine & dpath(i) & " - Not found"
	End if	
Next

'Copied available folders -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Endmessage = MsgBox(Conclusion,64,PGN)

'Finished backup ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ Backup done ---