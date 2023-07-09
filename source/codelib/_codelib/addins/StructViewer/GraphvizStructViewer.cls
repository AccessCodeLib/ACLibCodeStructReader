VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GraphvizStructViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: _codelib.addins.StructViewer.GraphvizStructViewer
'---------------------------------------------------------------------------------------
'
' GraphvizStructViewer
'
' Remarks:
'
'  @todo Setting of Graphviz parameters is still missing (there are still too many constants)
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/StructViewer/GraphvizStructViewer.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'  <use>_codelib/addins/shared/CodeModuleReader.cls</use>
'  <use>api/graphviz/GraphvizDOT.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const conDefaultMaxLevel As Long = 3

Private Const conDefaultGraphName As String = "DOT"

Private Const conDefaultFontname As String = "Verdana"

Private Const conObjGrpColor0 As String = "#0000FF" 'blau
Private Const conObjGrpColor1 As String = "#993399" 'lila
Private Const conObjGrpColor2 As String = "#996600" 'braun
Private Const conObjGrpColor3 As String = "#666600" 'grün

Private Const conObjGrpExtObjColor As String = "#006600" 'grün
Private Const conObjGrpExtObjColorOnlyInputOrOutput As String = "#FF0000" '"red"
Private Const conObjGrp_FirstGroupColor As String = "#000000"

Private Const conObjGrpStyleObj As String = ""
Private Const conObjGrpStyleObjGrp As String = "solid" '"dashed" '"bold"
Private Const conObjGrpFontnameFu As String = "Arial" '"Arial bold"
Private Const conDefaultObjGrpFontsizeFu As Long = 8 '10
Private Const conDefaultObjGrpFontsizeObjGrpDiff As Long = 2 'diese Zahlt wird zu Fu-Size addiert

Private Const conDOT_GraphProp As String = _
                  "graph [bgcolor=""#FFFFFF"", margin=0.05,overlap=false, splines=true, concentrade=true, fontname=""" & conDefaultFontname & """, ratio=compress];" & _
                  "node[color=""#000000"",fontname=""" & conDefaultFontname & """];" & _
                  "edge[color=""#000000"", arrowsize=0.8, labelangle=60, labelfontsize=" & conDefaultObjGrpFontsizeFu & ",labelfontname=""" & conDefaultFontname & """];"
               

Private m_DOT As GraphvizDOT

Private m_CodeModuleStructReader As CodeModuleStructReader

Private m_BaseCodeModuleName As String
Private m_RefreshStruct As Boolean
Private m_ShowProcedures As Boolean

Private Sub Class_Initialize()
   Set m_CodeModuleStructReader = New CodeModuleStructReader
End Sub

Public Property Get BaseCodeModuleName() As String

   BaseCodeModuleName = m_CodeModuleStructReader.BaseCodeModuleName
   
End Property

Public Property Let BaseCodeModuleName(ByVal CmName As String)

   m_CodeModuleStructReader.BaseCodeModuleName = CmName
   
End Property

Public Property Get ShowProcedures() As Boolean
   ShowProcedures = m_ShowProcedures
End Property

Public Property Let ShowProcedures(ByVal NewValue As Boolean)
   m_ShowProcedures = NewValue
End Property

Public Sub ReadStruct()

   m_CodeModuleStructReader.ReadStruct m_ShowProcedures
   
End Sub

Private Function ModulReaderExistsInCollection(Col As Collection, cmr As CodeModuleReader) As Boolean
   Dim TempModuleReader As CodeModuleReader
   For Each TempModuleReader In Col
      If TempModuleReader.Name = cmr.Name Then
         ModulReaderExistsInCollection = True
         Exit For
      End If
   Next
End Function

Public Sub CreateDOT( _
      Optional ByVal GraphMode As DOT_GraphMode = DOT_GraphMode.GraphMode_DOT, _
      Optional ByVal RankDir As DOT_RankDir = DOT_RankDir.RankDir_LR, _
      Optional ByVal GraphName As String = vbNullString, _
      Optional ByVal RndSeed As Long = 0, _
      Optional ByVal markFirstItem As Boolean = False)
   
   If m_CodeModuleStructReader.CodeModuleReaderCollection.Count = 0 Or m_RefreshStruct Then
      ReadStruct
   End If
   
   Set m_DOT = New GraphvizDOT
   
   m_DOT.Init GraphMode, RankDir, GraphName, , conDOT_GraphProp
   m_DOT.RndSeed = RndSeed
   
   AddDOTobjects markFirstItem
   m_DOT.ToBinaryGraph DOT2PNG, True
   
End Sub

Private Sub AddDOTobjects(ByVal markFirstItem As Boolean)
   
   Dim CmReader As CodeModuleReader
   
   Dim strNodeID As String
   Dim strColor As String
   Dim strStyle As String
   Dim CmSubGraph As GraphvizDOTgraph
   Dim DotNode As GraphvizDOTnode
   
   If markFirstItem Then
      strColor = "blue"
   End If
   
   For Each CmReader In m_CodeModuleStructReader.CodeModuleReaderCollection

      strNodeID = CmReader.Name
      If CmReader.Name = "modErrorHandler" Then
         strColor = "gray"
      Else
         strColor = vbNullString
      End If
      
      If m_ShowProcedures Then
         Set CmSubGraph = New GraphvizDOTgraph
         If CmReader.ComponentType = vbext_ct_StdModule Then
         If CmReader.Procedures.Count > 0 Then
            AddProcedureNodes CmReader, CmSubGraph
         End If
         End If
      End If
            
      m_DOT.AddNode strNodeID, CmReader.Name, , strColor, strStyle, , CmSubGraph, GetLink(CmReader)
      m_DOT.Node(strNodeID).LayoutMode = CmReader.CodeModule.Parent.Type

      AddDOTconnections CmReader, m_ShowProcedures
      
      Set CmSubGraph = Nothing
     
   Next
   
   If m_ShowProcedures Then
   For Each CmReader In m_CodeModuleStructReader.CodeModuleReaderCollection

      strNodeID = CmReader.Name
      Set DotNode = m_DOT.Nodes(strNodeID)

      If Not CmReader.ProcedureDependenciesCollection Is Nothing Then
      If CmReader.ProcedureDependenciesCollection.Count > 0 Then
         AddProceduresDOTconnections CmReader
      End If
      End If

   Next
   End If
   
End Sub

Private Sub AddProcedureNodes(ByVal CmReader As CodeModuleReader, ByVal SubGraph As GraphvizDOTgraph)
   
   Dim strNodeID As String
   Dim strColor As String
   Dim strStyle As String
   Dim Proc As CodeModuleProcedure
   
   For Each Proc In CmReader.Procedures
      strNodeID = CmReader.Name & "_" & Proc.Name
      SubGraph.AddNode strNodeID, Proc.Name
      'AddUsedInProcedureNodes Proc
   Next
   
End Sub

Private Sub AddUsedInProcedureNodes(ByVal ReferenceProc As CodeModuleProcedure)

   Dim strNodeID As String
   Dim Proc As CodeModuleProcedure
   Dim DotNode As GraphvizDOTnode
   
   If ReferenceProc.UsedInProcCollection Is Nothing Then
      Exit Sub
   End If
   
   For Each Proc In ReferenceProc.UsedInProcCollection
      strNodeID = Proc.ProcVbComponent.Name & "_" & Proc.Name
      Set DotNode = m_DOT.Nodes(Proc.ProcVbComponent.Name)
      If DotNode.SubGraph Is Nothing Then
         Set DotNode.SubGraph = New GraphvizDOTgraph
      End If
   On Error Resume Next
      DotNode.SubGraph.AddNode strNodeID, Proc.Name
   Next

End Sub

Private Function GetLink(cmr As CodeModuleReader) As String
   
   Dim strName As String
   Dim i As Long
   
   strName = Replace(cmr.Name, "_", "__")

   i = 2
   Do While i <= Len(strName)
      If StrComp(Mid(strName, i, 1), UCase(Mid(strName, i, 1)), vbBinaryCompare) = 0 _
         And StrComp(Mid(strName, i, 1), LCase(Mid(strName, i, 1)), vbBinaryCompare) <> 0 Then
         strName = Left(strName, i - 1) & "_" & LCase(Mid(strName, i, 1)) & Mid(strName, i + 1)
         i = i + 1
      End If
      i = i + 1
   Loop
   
   Select Case cmr.CodeModule.Parent.Type
      Case vbext_ct_StdModule
         strName = "namespace" & strName
      Case vbext_ct_ClassModule
         strName = "class_" & strName
      Case Else
         If Left(strName, 6) = "Form__" Then
            strName = Mid(strName, 7) & "_8frm"
         ElseIf Left(strName, 8) = "Report__" Then
            strName = Mid(strName, 9) & "_8rep"
         End If
         ' lassen
   End Select
   
   GetLink = LCase(strName) & ".html"
   
End Function

Private Sub AddDOTconnections(ByRef cmr As CodeModuleReader, Optional ByVal ClassesOnly As Boolean = False)

   Dim TempModuleReader As CodeModuleReader

   For Each TempModuleReader In cmr.RequiredModules
      If TempModuleReader.Name <> "modErrorHandler" Then  ' Errorhandler ignorieren
         If ClassesOnly Then
            If TempModuleReader.ComponentType = vbext_ct_ClassModule Then
               m_DOT.AddConnection TempModuleReader.Name, cmr.Name
            End If
         Else
            m_DOT.AddConnection TempModuleReader.Name, cmr.Name
         End If
      End If
   Next

End Sub

Private Sub AddProceduresDOTconnections(ByVal CmReader As CodeModuleReader)

   Dim ColItem As Object
   Dim Proc As CodeModuleProcedure
   Dim cmi As CodeModuleInfo
   
   For Each ColItem In CmReader.ProcedureDependenciesCollection
      If TypeOf ColItem Is CodeModuleProcedure Then
         Set Proc = ColItem
         AddProcedureDOTconnections Proc.ProcVbComponent.Name & "_" & Proc.Name, Proc.UsedInProcCollection
      Else
         Set cmi = ColItem
         AddProcedureDOTconnections cmi.Name, cmi.UsedInProcCollection
      End If
   Next

End Sub

Private Sub AddProcedureDOTconnections(ByVal LeftNodeId As String, ByVal ProcCollection As Collection)

   Dim Proc As CodeModuleProcedure
   Dim DotNode As GraphvizDOTnode

   If ProcCollection Is Nothing Then
      Exit Sub
   End If
   
   For Each Proc In ProcCollection
      If Proc.IsPrivate Or Proc.ProcVbComponent.Type <> vbext_ct_StdModule Then
         Set DotNode = m_DOT.Nodes(Proc.ProcVbComponent.Name)
   On Error Resume Next
         DotNode.SubGraph.AddNode Proc.ProcVbComponent.Name & "_" & Proc.Name, Proc.Name
      End If
   
      m_DOT.AddConnection LeftNodeId, Proc.ProcVbComponent.Name & "_" & Proc.Name
   Next

End Sub


'Private Function GetDotShape(cli As CodeLibInfo) As String
'
'   Select Case cli.Type
'      Case CodeLibElementType.clet_ClassModule
'         getDotShape = "octagon"
'      Case CodeLibElementType.clet_StdModule
'         getDotShape = "box"
'      Case CodeLibElementType.clet_Form
'         getDotShape = "octagon"
'      Case CodeLibElementType.clet_Report
'         getDotShape = "invhouse"
'   End Select
'
'End Function

'Private Function getDotColor(cli As CodeLibInfo) As String
'
'   Select Case cli.Type
'      Case CodeLibElementType.clet_Form
'         getDotColor = "#FFFFCC"
'      Case CodeLibElementType.clet_Report
'         getDotColor = "#FFFFCC"
'      Case Else
'         getDotColor = vbNullString
'   End Select
'
'End Function

'Private Function getDotStyle(cli As CodeLibInfo) As String
'
'   Select Case cli.Type
'      Case CodeLibElementType.clet_Form
'         getDotStyle = "filled"
'      Case CodeLibElementType.clet_Report
'         getDotStyle = "filled"
'      Case Else
'         getDotStyle = vbNullString
'   End Select
'
'End Function

#If EARLYBINDING Then
Public Property Set CurrentVbProject(ref As VBIDE.VBProject)
#Else
Public Property Set CurrentVbProject(ref As Object)
#End If

   Set m_CodeModuleStructReader.CurrentVbProject = ref

End Property

#If EARLYBINDING Then
Public Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Public Property Get CurrentVbProject() As Object
#End If

   Set CurrentVbProject = m_CodeModuleStructReader.CurrentVbProject

End Property
