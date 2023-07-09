VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: CodeModuleInfo
'---------------------------------------------------------------------------------------
'
' Auxiliary class for processing a CodeModule
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/CodeModuleInfo.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Name As String

Public ComponentType As vbext_ComponentType

Public UsedInProcCollection As Collection
