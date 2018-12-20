VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Blank with all subs or functions for the purpose of dealing with blank cells

Sub ReplaceBlanks()
'
' Replace blanks
'
    Range("A1").Select
    If ActiveCell.Offset(1, 0) <> "-" Then ' If did not replace before
        Range(ActiveCell, ActiveCell.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
        Range(Selection, Selection.End(xlDown)).Select
    Else ' If replaced before
        Range(ActiveCell, ActiveCell.End(xlToRight)).Select
        Range(Selection, Selection.End(xlDown)).Select
    End If
    Replace
    
End Sub

Private Sub Replace()
'
' Replace blanks with "-"
'
    Cells.Replace What:="", Replacement:="-", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

Private Sub ReplaceBlank(c As Range)
'
' This Sub will be called after the initial blank cells replacement
'
    c.EntireColumn.End(xlUp).Select ' Reset to the first cell of the column that has the selected cell
    Dim i, j, rowLoc, sizeRow, sizeCol As Integer
    
    rowLoc = Range(c, c.EntireRow.End(xlToLeft)).Columns.count ' calculate the distance of the current column to the first column in the row
    sizeRow = noRow(c.Offset(0, -rowLoc + 1)) ' call function to calculate how many columns are in the file
    sizeCol = noCol(c.Offset(0, -rowLoc + 1)) ' call function to calculate how many rows are in the file
    For j = 0 To sizeRow
        For i = 0 To sizeCol
        ' Replace the blank cells line by line, one by one
            If c.Offset(j, -rowLoc + i + 1).Value = "" Then
                c.Offset(j, -rowLoc + i + 1).Value = "-"
            End If
        Next i
    Next j

End Sub

Private Function noRow(c As Range) As Integer
'
' Count the number of rows in a file
'
    Dim sizeRow As Integer
    sizeRow = Range(c, c.EntireColumn.End(xlDown)).Rows.count
    noRow = sizeRow
End Function

Private Function noCol(c As Range) As Integer
'
' Count the number of columns in a file
'
    Dim sizeCol As Integer
    sizeCol = Range(c, c.EntireRow.End(xlToRight)).Columns.count
    noCol = sizeCol
End Function

Sub ReplaceBack()
'
' Replace "-" back to blanks
'
' We clear the "-" we filled in blanks to its original blanks
' This is for the future importing to other softwares
    Cells.Replace What:="-", Replacement:="", LookAt:=xlPart, SearchOrder:= _
            xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End Sub

