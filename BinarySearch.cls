Option Explicit

Private Function BinarySearchRow( _
    ByRef wsLookup As Worksheet _
  , ByVal rowFirst _
  , ByVal strSearchArg1 As String _
  , ByVal idxCol1 As Long _
  , Optional ByVal strSearchArg2 As String _
  , Optional ByVal idxCol2 As Long _
  , Optional ByVal strSearchArg3 As String _
  , Optional ByVal idxCol3 As Long) As Long
  
    Dim rowLast As Long
    Dim strSearchArg As String
    
    rowLast = FindLastRow(wsLookup)
    strSearchArg = strSearchArg1 & strSearchArg2 & strSearchArg3
    
    Do While True
        Dim rowMiddle As Long
        Dim strValue As String
        If rowFirst > rowLast Then
            Exit Do
        End If
        rowMiddle = Round((rowLast - rowFirst) / 2 + rowFirst)
        
        With wsLookup
        strValue = .Cells(rowMiddle, idxCol1)
        If idxCol2 <> 0 Then strValue = strValue & .Cells(rowMiddle, idxCol2)
        If idxCol3 <> 0 Then strValue = strValue & .Cells(rowMiddle, idxCol3)
        End With
        
        If CDbl(strSearchArg) < CDbl(strValue) Then
            rowLast = rowMiddle - 1
        ElseIf CDbl(strSearchArg) > CDbl(strValue) Then
            rowFirst = rowMiddle + 1
        Else
            BinarySearchRow = rowMiddle
            Exit Do
        End If
    Loop
End Function

Public Sub SortForBinary( _
    ByRef wsLookup As Worksheet _
  , ByVal idxCol1 As Long _
  , Optional ByVal idxCol2 As Long _
  , Optional ByVal idxCol3 As Long _
  , Optional ByVal rowHeader As Long = 1 _
  , Optional ByVal colFirst As Long = 1)
  
    Dim i As Long: i = rowHeader
    dim idxColMerged as long
    idxColMerged = FindLastColumn(wsLookup, rowHeader) + 1

    Do While wsLookup.Cells(i, colFirst) <> vbNullString
        Cells(i, idxColMerged) _
            = wsLookup.Cells(i, idxCol1) _
            & wsLookup.Cells(i, idxCol2) _
            & wsLookup.Cells(i, idxCol3)
            i = i + 1
    Loop
    ws.Range(ws.Cells(rowHeader, colFirst) _
           , ws.Cells(FindLastRow(ws), FindLastColumn(ws, rowHeader))).Sort _
        Key1:=ws.Cells(rowHeader, idxColMerged), _
        Order1:=xlAscending, _
        Header:=xlYes, _
        DataOption1:=xlSortTextAsNumbers
    ws.Columns(idxColMerged).ClearContents
End Sub
