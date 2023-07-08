const AddInName = "ACLib-StructReader"
const AddInFileName = "ACLibStructReader.accda"
const MsgBoxTitle = "Install ACLib-StructReader"

MsgBox "Before updating the add-in file, the add-in must not be loaded!" & chr(13) & _
       "Close all access instances for safety.", , MsgBoxTitle & ": Hinweis"

Select Case MsgBox("Should the add-in be used as a compiled file (accde)?" + chr(13) & _
                   "(Add-In is compiled and copied to the Add-In directory.)", 3, MsgBoxTitle)
   case 6 ' vbYes
      CreateMde GetSourceFileFullName, GetDestFileFullName
      MsgBox "Add-In was compiled and saved in '" + GetAddInLocation + "'.", , MsgBoxTitle
   case 7 ' vbNo
      FileCopy GetSourceFileFullName, GetDestFileFullName
      MsgBox "Add-In was saved in '" + GetAddInLocation + "'.", , MsgBoxTitle
   case else
      
End Select


'##################################################
' Utility functions:

Function GetSourceFileFullName()
   GetSourceFileFullName = GetScriptLocation & AddInFileName 
End Function

Function GetDestFileFullName()
   GetDestFileFullName = GetAddInLocation & AddInFileName 
End Function

Function GetScriptLocation()
   With WScript
      GetScriptLocation = Replace(.ScriptFullName & ":", .ScriptName & ":", "") 
   End With
End Function

Function GetAddInLocation()
   GetAddInLocation = GetAppDataLocation & "Microsoft\AddIns\"
End Function

Function GetAppDataLocation()
   Set wsShell = CreateObject("WScript.Shell")
   GetAppDataLocation = wsShell.ExpandEnvironmentStrings("%APPDATA%") & "\"
End Function

Function FileCopy(SourceFilePath, DestFilePath)
   set fso = CreateObject("Scripting.FileSystemObject") 
   fso.CopyFile SourceFilePath, DestFilePath
End Function

Function CreateMde(SourceFilePath, DestFilePath)
   Set AccessApp = CreateObject("Access.Application")
   AccessApp.SysCmd 603, (SourceFilePath), (DestFilePath)
End Function
