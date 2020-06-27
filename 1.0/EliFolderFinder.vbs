PGN = "EliFolderFinder"
Set objShellApp  = CreateObject( "Shell.Application" )

Function SelectFolder( myStartFolder )
    Set objFolder = objShellApp.BrowseForFolder( 0, "Backup-Ordner auswählen:", 0, myStartFolder )
    SelectFolder = objFolder.Self.Path
End Function

backuppath = SelectFolder( "" )
Ausg = MsgBox(backuppath)