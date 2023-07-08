Attribute VB_Name = "modGraphviz"
Option Compare Database
Option Explicit

Private Const p_conWordWrapMinLen As Long = 16


'Public Property Get GraphvizSetting(sProp As String) As String
'   GraphvizSetting = Nz(CLookupSQL("Select PropertyValue FROM ADF_GraphvizProperties WHERE PropertyDef='" & sProp & "'"), vbNullString)
'End Property
'
'Public Property Let GraphvizSetting(sProp As String, sNewValue As String)
'   Dim strSQL As String
'   If CLookupSQL("SELECT COUNT(*) FROM ADF_GraphvizProperties WHERE PropertyDef='" & sProp & "'") > 0 Then
'      strSQL = "UPDATE ADF_GraphvizProperties SET PropertyValue='" & sNewValue & "' WHERE PropertyDef='" & sProp & "'"
'   Else
'      strSQL = "INSERT INTO ADF_GraphvizProperties (PropertyDef,PropertyValue) VALUES ('" & sProp & "', '" & sNewValue & "')"
'   End If
'   CodeDb.Execute strSQL
'End Property
