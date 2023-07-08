Attribute VB_Name = "___TEST"
Option Compare Database
Option Explicit


Private Sub Test()

#If EARLYBINDING Then
   Dim vbp As VBIDE.VBProject
#Else
   Dim vbp As Object
#End If
   Dim cmr As CodeModuleReader
   
   Set vbp = VBE.ActiveVBProject
   Set cmr = New CodeModuleReader
   
   Set cmr.CodeModule = vbp.VBComponents("modWinAPI").CodeModule
   
   Dim cnt As Long
   cnt = cmr.CheckDependency(vbp)
   
   Dim TempModuleReader As CodeModuleReader
   For Each TempModuleReader In cmr.RequiredModules
      Debug.Print TempModuleReader.Name
   Next

End Sub

Private Sub CodeModuleStructReader_All()
   
   Dim structReader As CodeModuleStructReader
   Set structReader = New CodeModuleStructReader
   
   'structReader.CreateDOT DOT_GraphMode.GraphMode_NEATO

End Sub


Private Sub testGetLink()

   Debug.Print GetLink("abcXYZ")

End Sub

Private Function GetLink(strName As String) As String
   

   Dim i As Long

   i = 2
   Do While i <= Len(strName)
      If StrComp(Mid(strName, i, 1), UCase(Mid(strName, i, 1)), vbBinaryCompare) = 0 _
         And StrComp(Mid(strName, i, 1), LCase(Mid(strName, i, 1)), vbBinaryCompare) <> 0 Then
         strName = Left(strName, i - 1) & "_" & LCase(Mid(strName, i, 1)) & Mid(strName, i + 1)
         i = i + 1
      End If
      i = i + 1
   Loop
   
   GetLink = strName

End Function
