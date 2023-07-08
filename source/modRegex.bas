Attribute VB_Name = "modRegex"
Option Compare Database
Option Explicit

Private m_Regex As Object

Public Property Get RegEx() As Object
   If m_Regex Is Nothing Then
      Set m_Regex = CreateObject("Vbscript.Regexp")
   End If
   Set RegEx = m_Regex
End Property
