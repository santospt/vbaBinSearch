Option Explicit
Public Function FindLastRow(Optional sht As Worksheet, Optional idxCol As Long = 1, Optional fromTop As Boolean = False) As Long
  
    Dim firstRow As Long
  
    If sht Is Nothing Then Set sht = ActiveSheet
    
    If fromTop Then
        If sht.Cells(1, idxCol) <> vbNullString Then
            firstRow = 1
        Else
            firstRow = sht.Cells(1, idxCol).End(xlDown).Row
        End If
        
        If sht.Cells(firstRow + 1, idxCol) = vbNullString Then
            FindLastRow = firstRow
        Else
            FindLastRow = sht.Cells(firstRow, idxCol).End(xlDown).Row
        End If
    Else
        FindLastRow = sht.Cells(sht.Rows.Count, idxCol).End(xlUp).Row
    End If
    
End Function

Public Function FindLastColumn(Optional sht As Worksheet, Optional idxRow As Long = 1, Optional fromLeft As Boolean = False) As Long
  
    Dim firstCol As Long
    
    If sht Is Nothing Then Set sht = ActiveSheet
    
    If fromLeft Then
        If sht.Cells(idxRow, 1) <> vbNullString Then
            firstCol = 1
        Else
            firstCol = sht.Cells(idxRow, 1).End(xlToRight).Column
        End If
        
        If sht.Cells(idxRow, firstCol + 1) = vbNullString Then
            FindLastColumn = firstCol
        Else
            FindLastColumn = sht.Cells(idxRow, firstCol).End(xlToRight).Column
        End If
    Else
        FindLastColumn = sht.Cells(idxRow, sht.Columns.Count).End(xlToLeft).Column
    End If
    
End Function
