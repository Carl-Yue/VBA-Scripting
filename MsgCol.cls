VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MsgCol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class MsgCol with all subs or functions for the purpose of massage columns if the columns need to be processed

Private Sub MasgCol()
' Massage a column if it is not standard
    ActiveCell.EntireColumn.End(xlUp).Select
    ReplaceBlank ActiveCell
    SplitCol
    CombineCol
End Sub

Sub SplitCol()
'
' Separate Columns when necessary
'
    InsertCol
    ActiveCell.Offset(2, 0).Select
    SplitBy
    ActiveCell.EntireColumn.End(xlUp).Select

End Sub

Private Sub SplitBy()
'
' Separate Columns by ","
'
    Range(Selection, Selection.End(xlDown)).Select
    Selection.TextToColumns Destination:=ActiveCell, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
End Sub
Private Function DateToText(c As Range) As String
' Transfer Date data to Text data

    DateToText = Format(c.Value, "yyyy/mm/dd")
End Function
Private Function TimeToText(c As Range) As String
' Transfer Time data to Text data

    TimeToText = Format(c.Value, "hh:mm")
End Function

Sub CombineCol()
' Combine 2 columns
    Dim i, j As Integer
    Dim sizeRow As Integer
    sizeRow = noRow(ActiveCell)
    ' Wherever curser is, move to the top of this column
    ActiveCell.EntireColumn.End(xlUp).Select
    ' Combine row by row
    hs = CountHeader()
    
    MsgBox hs
    ' Combine header firstly
    For j = 0 To hs - 1
        ActiveCell.Offset(j, 0).Value = ActiveCell.Offset(j, 0).Value & ActiveCell.Offset(j, 1).Value
    Next j
    ' Combine the data
    For i = hs + 1 To sizeRow
        ActiveCell.Offset(i, 0).Value = DateToText(ActiveCell.Offset(i, 0)) & " " & TimeToText(ActiveCell.Offset(i, 1))
    Next i
    ' Delete the redundant column
    DeleteCol ActiveCell.Offset(0, 1)
End Sub

