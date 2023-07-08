Attribute VB_Name = "_initWizard"
'---------------------------------------------------------------------------------------
' Modul: _initWizard
'---------------------------------------------------------------------------------------
'/* *
' <summary>
' Initialisierungsaufruf des Add-Ins
' </summary>
' <remarks></remarks>
'* */
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'---------------------------
' Initialisierungsfunktion
'---------------------------
Private Function initWizard() As Boolean
   If Application.CurrentDb Is Nothing Then
      MsgBox "Bitte öffnen Sie zuerst eine Access-Anwendung.", vbCritical
      initWizard = False
      Exit Function
   End If
   initWizard = StartApplication
End Function

Public Function StartWizard() As Variant
   If Not initWizard Then Exit Function
End Function
