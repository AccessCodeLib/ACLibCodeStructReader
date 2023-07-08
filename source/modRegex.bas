Attribute VB_Name = "modRegex"
Option Compare Database
Option Explicit

Private m_RegExp As Object

Public Property Get RegExp() As Object
   If m_RegExp Is Nothing Then
      Set m_RegExp = CreateObject("VbScript.RegExp")
   End If
   Set RegExp = m_RegExp
End Property
