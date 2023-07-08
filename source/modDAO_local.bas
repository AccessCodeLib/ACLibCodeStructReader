Attribute VB_Name = "modDAO_local"
Option Compare Database
Option Explicit

Public Function CLookupSQL(sSQL As String) As Variant

    Dim rst As DAO.Recordset

    Set rst = CodeDb.OpenRecordset(sSQL, dbOpenForwardOnly)
    With rst
        If .EOF Then
            CLookupSQL = Null
        Else
            CLookupSQL = .Fields(0)
        End If
        .Close
    End With
    Set rst = Nothing

End Function

Public Function openLokalRecordset(ByVal Source As String, _
                        Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenForwardOnly, _
                        Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                        Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset

On Error GoTo HandleErr
'
'    If (RecordsetOptions And dbSeeChanges) = 0 Then
'        RecordsetOptions = RecordsetOptions + dbSeeChanges
'    End If
    Set openLokalRecordset = CodeDb.OpenRecordset(Source, RecordsetType, RecordsetOptions, LockEdit)
ExitHere:
   Exit Function

HandleErr:
   Err.Raise Err.Number, "openDAORecordset:" & Err.Source, Err.Description
   Resume ExitHere

End Function
