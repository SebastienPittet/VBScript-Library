' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨

'    Author: sebastien at pittet dot org
'    Date  :
'    Goal  : Provide some cool functions...
'  Version : Library.vbs v.1.3

' ¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨¨°º©o¿,,¿o©º°¨

On Error Resume Next

'@@@@@@@@@@@@@@@
'Functions & Sub
'@@@@@@@@@@@@@@@

Function UDate(oldDate)
  'Determine the Epoch Time
  UDate = DateDiff("s", "01/01/1970 00:00:00", oldDate)
End Function
'-------------------------------------------------------------------
Function unUDate(intTimeStamp)
  'Reverse Epoch Time
  unUDate = DateAdd("s", intTimeStamp, "01/01/1970 00:00:00")
end Function
'-------------------------------------------------------------------

Sub RemoveThatScriptAtNextReboot
      Dim WshShell
      Dim fso
      Dim ScriptName
      
      Set WshShell = WScript.CreateObject("WScript.Shell")
      Set fso = CreateObject("Scripting.FileSystemObject")
      ScriptName=fso.GetFile(wscript.scriptfullname)  	

      WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce\RemoveScript", "CMD.EXE /C ""DEL " & ScriptName & """"
End Sub
'-------------------------------------------------------------------

'Delete a File
  Sub DeleteFile(FileFullPath)
     Set fso = CreateObject("Scripting.FileSystemObject")
     If fso.FileExists(FileFullPath) Then
        fso.DeleteFile(FileFullPath)
     End If
  End Sub
'-------------------------------------------------------------------

'Ecriture dans un fichier texte
   Sub AppendLineToFile(MonFichier, MaLigne)
      Const ForReading = 1, ForWriting = 2, ForAppending = 8
      Dim fso, f
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set f = fso.OpenTextFile(MonFichier, ForAppending, True)
      f.WriteLine MaLigne
   End Sub
'-------------------------------------------------------------------

'Récupération du domaine, Computer, Username
   Function GetDomain
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      GetDomain = WshNetwork.UserDomain
   End Function
   
	Function GetLDAPDomain
		'Get Domain name from RootDSE object.
		'Exemple : DC=domain,DC=ch
		Set objRootDSE = GetObject("LDAP://RootDSE")
		GetLDAPDomain = objRootDSE.Get("DefaultNamingContext")
	End Function

   Function GetComputerName
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      GetComputerName = WshNetwork.ComputerName
   End Function

   Function GetUserName
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      GetUserName = WshNetwork.UserName
   End Function
   
   Function GetSite(HostName)
       'Retourne le nom du site ADS où le PC se trouve
       'Requis : IADSTools.dll enregistrée (support tools)
       'le paramètre Hostname peut être fourni par GetComputerName()
       Dim objIadsTools
	   Set objIadsTools = CreateObject("IADsTolls.DCFunctions")
	   GetSite = objIadsTools.DsGetSiteName(HostName)
   End Function
   
'--------------------------------------------------------------------
   
'Retourne la version de l'Operating System
Function GetOSVersion(Computername)
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\" & Computername & "\root\cimv2")
	Set colOperatingSystems = objWMIService.ExecQuery _
	    ("Select * from Win32_OperatingSystem")
	For Each objOperatingSystem in colOperatingSystems
	    GetOSVersion = objOperatingSystem.Caption & " " & _
	        objOperatingSystem.Version
	Next
End Function

'--------------------------------------------------------------------

'Delete a shortcut into a Special Folder
   Sub DeleteSpecialShortcut(SpecialLocation, ShortcutName)
      'Special Folders are :
	'AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup
	'Desktop, Favorites, Fonts, MyDocuments, NetHood, PrintHood, Programs
	'Recent, SendTo, StartMenu, Startup, Templates.

      Set Shell = CreateObject("WScript.Shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      ShortcutPath = Shell.SpecialFolders(SpecialLocation) & "\" & ShortcutName & ".lnk"
      If fso.FileExists(ShortcutPath) Then
         FSO.DeleteFile ShortcutPath
      End If
   End Sub
'--------------------------------------------------------------------

'Delete a Directory into a Special Folder
   Sub DeleteSpecialDirectory(SpecialLocation, DirectoryName)
      'Special Folders are :
      'AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup
      'Desktop, Favorites, Fonts, MyDocuments, NetHood, PrintHood, Programs
      'Recent, SendTo, StartMenu, Startup, Templates.

      Set Shell = CreateObject("Wscript.shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      DirectoryPath = Shell.SpecialFolders(SpecialLocation) & "\" & DirectoryName
      If fso.FolderExists(DirectoryPath) Then
         FSO.DeleteFolder DirectoryPath
      End If
   End Sub
'--------------------------------------------------------------------

'Move a Directory into a Special Folder
   Sub MoveSpecialDirectory(SpecialLocationSource, DirectoryNameSource, SpecialLocationDestination, DirectoryNameDestination)
      'Special Folders are :
      'AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup
      'Desktop, Favorites, Fonts, MyDocuments, NetHood, PrintHood, Programs
      'Recent, SendTo, StartMenu, Startup, Templates.

      Set Shell = CreateObject("Wscript.shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      
      SourceDirectoryPath = Shell.SpecialFolders(SpecialLocationSource) & "\" & DirectoryNameSource
      DestinationDirectoryPath = Shell.SpecialFolders(SpecialLocationDestination) & "\" & DirectoryNameDestination      
   
      If Not fso.FolderExists(SourceDirectoryPath) Then
         fso.MoveFolder SourceDirectoryPath , DestinationDirectoryPath
      End If
   End Sub
'--------------------------------------------------------------------

'Move a shortcut
   Sub MoveShortcut(SpecialLocationSource, ShortcutName, SpecialLocationDestination, PathDestination)
      'Special Folders are :
      'AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup
      'Desktop, Favorites, Fonts, MyDocuments, NetHood, PrintHood, Programs
      'Recent, SendTo, StartMenu, Startup, Templates.

      Set Shell = CreateObject("Wscript.shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      
      ShortcutPath = Shell.SpecialFolders(SpecialLocationSource) & "\" & ShortcutName & ".lnk"
      Destination = Shell. SpecialFolders(SpecialLocationDestination) & "\" & PathDestination & "\"
   
      If fso.FileExists(ShortcutPath) Then
         fso.MoveFile ShortcutPath , Destination
      End If
   End Sub
'--------------------------------------------------------------------

'Read a Registry Key
   Function RegRead(RegKey)
      Dim WshShell
      Set WshShell = WScript.CreateObject("WScript.Shell")
      RegRead = WshShell.RegRead(RegKey)
   End Function
'--------------------------------------------------------------------

'Delete a Registry Key
  Sub RegDelete(RegValue)
  	Dim WshShell
  	Set WshShell = WScript.CreateObject("Wscript.Shell")
  	WshShell.RegDelete(RegValue)
  	WshShell = Nothing
  End Sub

'--------------------------------------------------------------------

'Write a Registry Key
   Sub RegWrite(RegKey, Value)
      Dim WshShell
      Set WshShell = WScript.CreateObject("WScript.Shell")
      WshShell.RegWrite RegKey, Value
   End Sub
'--------------------------------------------------------------------

'Determine the Program Files Dir(usefull for German's Workstations)
   Function GetProgFilesDir
      Dim WshShell
      Set WshShell = WScript.CreateObject("WScript.Shell")
      GetProgFilesDir = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\ProgramFilesDir")
   End Function
'--------------------------------------------------------------------

'Remove DriveMap
   Sub RemoveDriveMap(Lettre)
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      Set fso = CreateObject("Scripting.FileSystemObject")
      If fso.DriveExists(Lettre) Then
          WshNetwork.RemoveNetworkDrive Lettre, True
      End If
   End Sub
'--------------------------------------------------------------------

'Create DriveMap
   'Lettre = H:  et Chemin = \\server\share
   Sub CreateDriveMap(Lettre, Chemin)
      Set WshNetwork = WScript.CreateObject("WScript.Network")
      Set fso = CreateObject("Scripting.FileSystemObject")
      If fso.DriveExists(Lettre) Then
        WshNetwork.RemoveNetworkDrive Lettre, True
      End If
      WshNetwork.MapNetworkDrive Lettre, Chemin
   End Sub
'--------------------------------------------------------------------

'Ouverture d'un fichier texte dans un directory donné
    Sub CreateTextFile(FileFullPath)
      Dim fso, MyFile
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set MyFile = fso.CreateTextFile(FileFullPath, True)
      MyFile.Close
    End Sub
'--------------------------------------------------------------------

'Exécution d'une commande shell
   Sub ShellRun(CommandLine)
      Set WshShell = WScript.CreateObject("WScript.Shell")
      WshShell.Run CommandLine,2,True
   End Sub
'--------------------------------------------------------------------

'Create Folder
   Sub CreateFolder(FolderFullPath)
      Set fso = CreateObject("Scripting.FileSystemObject")
      If Not fso.FolderExists(FolderFullPath) Then
         Set FolderObject = fso.CreateFolder(FolderFullPath)
      End If
   End Sub
'--------------------------------------------------------------------

'Check If File Exist at a specified path (Boolean)
   Function FileExist (FileFullPath)
      Dim Fso
      Set Fso = CreateObject("Scripting.FileSystemObject")
      If (Fso.FileExists(FileFullPath)) Then
         FileExist = True
      Else
         FileExist = False
      End If
   End Function
'--------------------------------------------------------------------

'Check if a Specified Folder Exist (Boolean)
   Function FolderExist(fldr)
      Dim fso, msg
      Set fso = CreateObject("Scripting.FileSystemObject")
      If (fso.FolderExists(fldr)) Then
         status = True
      Else
         Status = False
      End If
      FolderExist = status
   End Function
'--------------------------------------------------------------------

  Function FolderSize(FolderPath)  ' return folder size
    Dim fso
	Dim fsoFolder
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsoFolder = fso.GetFolder(FolderPath)
	
	FolderSize = fsoFolder.Size & " bytes"
	
	Set fsoFolder = Nothing
    Set fso = Nothing
  End Function
'--------------------------------------------------------------------



'Rend l'Operating system, grâce au Build Number
   Function GetOSBuildNumber
      Const BUILD_WINDOWS_NT = "1381"
      Const BUILD_WINDOWS_2000 = "2195"
      Const BUILD_WINDOWS_XP = "2600"

      On Error Resume Next
      OSBuildNumber = "Unknown"

      Dim objShell
      Dim strBuildNumber

      Set WshShell = CreateObject("WScript.Shell")
      GetOSBuildNumber = WshShell.RegRead("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentBuildNumber")
   End Function
'--------------------------------------------------------------------

'Retourne la valeur d'une variable d'environnement
   Function GetSystemVariable(Variable)
      'Taper "set" dans le prompt MS-DOS pour obtenir la liste des variables
      Set WshShell = WScript.CreateObject( "WScript.Shell" )
      GetSystemVariable = WshShell.ExpandEnvironmentStrings("%" & Variable & "%")
   End Function
'--------------------------------------------------------------------

'Ajoute un event dans l'event log
   Sub WriteLogEvent(EventType, EventText)
      Const SUCCESS = 0
      Const Error = 1
      Const WARNING = 2
      Const INFORMATION = 4
      Const AUDIT_SUCCESS = 8
      Const AUDIT_FAILURE = 16

      Set WshShell = WScript.CreateObject("WScript.Shell")
      WshShell.LogEvent EventType, EventText
   End Sub
'--------------------------------------------------------------------

'Retourne le chemin du dossier d'où le script est exécuté
   Function GetThisFolderPath()
      Set fso = CreateObject("Scripting.FileSystemObject")
      Set file = fso.GetFile(wscript.scriptfullname)
      GetThisFolderPath=File.ParentFolder
   End Function
'--------------------------------------------------------------------

'Retourne le nom du script
   Function GetScriptName()
      Set fso = CreateObject("Scripting.FileSystemObject")
      GetScriptName=fso.GetFile(wscript.scriptfullname)  	
   End  Function
'--------------------------------------------------------------------

'Returns the OS Language
   Function GetOSLanguage
		Dim languageNR 'String, contains the language number
		
		strComputer = "."
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		Set colOperatingSystems = objWMIService.ExecQuery ("SELECT OSLanguage FROM Win32_OperatingSystem")
		
		For Each objOperatingSystem in colOperatingSystems
		  LanguageNR = objOperatingSystem.oslanguage
		Next
		
		Select Case languageNR
		  Case 1036
		    GetOSLanguage = "French"
		  Case 1031
		    GetOSLanguage = "German"
		  Case 0409 
		    GetOSLanguage = "English"
		  Case Else
		    GetOSLanguage = "Error"
		End Select
	End Function
'--------------------------------------------------------------------

'Create a Folder in the Start Menu
  Sub CreateSpecialFolder(SpecialLocation, FolderName)
      'Special Folders are :
      'AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup
      'Desktop, Favorites, Fonts, MyDocuments, NetHood, PrintHood, Programs
      'Recent, SendTo, StartMenu, Startup, Templates.
      Set Shell = CreateObject("Wscript.shell")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      FolderPath = Shell.SpecialFolders(SpecialLocation) & "\" & FolderName
      
      If Not fso.FolderExists(FolderPath) Then
         Set FolderObject = fso.CreateFolder(FolderPath)
      End If
  End Sub
'--------------------------------------------------------------------

'Retrieve File Version
  Function GetFileVersion(FileFullPath)
     Set objFSO = CreateObject("Scripting.FileSystemObject")
     GetFileVersion = objFSO.GetFileVersion(FileFullPath)
  End Function
'-----------------------------------------------------------------------   

'Retrieve Client Access Version in the format "VRM" => VersionReleaseModification (ex : 510)
  Function GetIBMCAVersion()
      On Error Resume Next
      Dim WshShell
      Dim Version
      Dim Release
      Dim Modification
      
      Set WshShell = WScript.CreateObject("WScript.Shell")
      Version = WshShell.RegRead("HKLM\Software\IBM\Client Access\CurrentVersion\Version")
      Release = WshShell.RegRead("HKLM\Software\IBM\Client Access\CurrentVersion\Release")
      Modification = WshShell.RegRead("HKLM\Software\IBM\Client Access\CurrentVersion\ModificationLevel")
      
      GetIBMCAVersion = Version & Release & Modification
  End Function
'-----------------------------------------------------------------------   

' Function to check if a process is running
function isProcessRunning(byval strComputer,byval strProcessName)

	Dim objWMIService, strWMIQuery

	strWMIQuery = "Select * from Win32_Process where name like '" & strProcessName & "'"
	
	Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _ 
			& strComputer & "\root\cimv2") 

	if objWMIService.ExecQuery(strWMIQuery).Count > 0 then
		isProcessRunning = "true"
	else
		isProcessRunning = "false"
	end if

end function


'@@@@@@@@@@@@@
' MAIN PROGRAM
'@@@@@@@@@@@@@

wscript.echo "Ce fichier n'est pas conçu pour être exécuté..." &vbCrLf & "Edition requise."
WScript.Echo "Today is : " & Now()


