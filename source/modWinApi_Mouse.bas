Attribute VB_Name = "modWinApi_Mouse"
Option Compare Database
Option Explicit

' aus http://www.mvps.org/access/api/api0044.htm

' This code was originally written by Terry Kreft.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Terry Kreft
'

Public Enum IDC_MouseCursor
   IDC_APPSTARTING = 32650&
   IDC_HAND = 32649&
   IDC_ARROW = 32512&
   IDC_CROSS = 32515&
   IDC_IBEAM = 32513&
   IDC_ICON = 32641&
   IDC_no = 32648&
   IDC_Size = 32640&
   IDC_SIZEALL = 32646&
   IDC_SIZENESW = 32643&
   IDC_SIZENS = 32645&
   IDC_SIZENWSE = 32642&
   IDC_SIZEWE = 32644&
   IDC_UPARROW = 32516&
   IDC_WAIT = 32514&
End Enum

Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" _
   (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

'Declare Function LoadCursorFromFile Lib "user32" Alias _
'   "LoadCursorFromFileA" (ByVal lpFileName As String) As Long

Declare Function SetCursor Lib "user32" _
   (ByVal hCursor As Long) As Long

Function MouseCursor(CursorType As IDC_MouseCursor)
  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Function

'Function PointM(strPathToCursor As String)
'  Dim lngRet As Long
'  lngRet = LoadCursorFromFile(strPathToCursor)
'  lngRet = SetCursor(lngRet)
'End Function
