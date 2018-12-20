Attribute VB_Name = "main"
Sub main()
'
' There are 6 steps when importing documents for iSeries and Tendam accounts
' The importing process depends on the source of data collection
'
    Dim im As New Import
    Dim bl As New Blank
    Dim ws As Worksheet
    im.ImportFiles ' Step 1
    ' The sub is in Import class
    ' Import files depends on if it is delimited or fixed fields
    For Each ws In ThisWorkbook.Worksheets
    ' loop through all the worksheets in workbook
        ws.Activate
        bl.ReplaceBlanks ' Step 2
        ' Some of cells are empty, it will hinder our way of processing, so we replace the blank cells to "-"
        ReplaceEqual ' Step 3
        ' We need to pre-process the header rows, and put the end of header rows to "="
        CombineHeaderAuto ' Step 4
        ' The header will take several rows, so we need to combine them together
        DeleteEmptyRow ' Step 5
        ' We will delete the blank rows
        bl.ReplaceBack ' Step 6
    Next
End Sub

Sub ReplaceEqual()
'
' Step 3: Replace the row under header with "="
'
    Range("A1").Select ' Reset the selection to the first cell
    Dim counter, sizeCol As Integer
    Dim c As Range
    counter = 1
    Do Until ActiveCell.EntireRow.End(xlToLeft).Offset(counter, 0) <> "-"
    ' loop counts how many rows in the header
        counter = counter + 1
    Loop
    sizeCol = noCol(ActiveCell) ' call function to calculate how many columns are in the file
    For Each c In Range(ActiveCell.Offset(counter, 0), ActiveCell.Offset(counter, 0).EntireRow.End(xlToRight))
    ' Replace the row under the header with "="
        c.Value = "="
    Next
End Sub



Sub CombineHeaderAuto()
'
' Step 4: Automatically combine the header together
'
    Range("A1").Select ' Reset the selection to the first cell
    Dim counter As Integer
    counter = 1
    Do While ActiveCell.Offset(0, counter).Value <> ""
        If ActiveCell.Offset(1, counter).Value <> "-" Then
        ' Combine the header rows if they were not blank
            CombineHeader ActiveCell.Offset(0, counter)
        End If
        counter = counter + 1
    Loop
End Sub

Sub CombineHeader(c As Range)
'
' Combine header for a column
'
    c.EntireColumn.End(xlUp).Select ' Reset the selection to the top of current column
    Dim counter, i, j As Integer
    counter = CountHeader()
    ' Put the content together
    For i = 1 To counter - 1
        c.Value = c.Value & " " & c.Offset(i, 0).Value
    Next i
    ' Clear the redundent info
    For j = 1 To counter - 1
        c.Offset(j, 0).Value = "-"
    Next j
    Range("A1").Select
    
End Sub

Sub DeleteEmptyRow()
'
' Step 5: Delete the Empty Rows
'
    DeleteBlankRow "-" ' Delete rows if all cells in the row are "-"
    DeleteBlankRow "=" ' Delete rows if all cells in the row are "="
End Sub

Private Sub DeleteBlankRow(x As String)
'
' Delete rows if the cells in the row are all with same specified character
'
    Dim i, j, delCount, sizeCol, sizeRow As Integer
    Dim result As Boolean
    
    Range("A1").Select ' Reset the selection to the first cell
    sizeRow = noRow(ActiveCell) ' Calculate the number of rows in the spreadsheet
    sizeCol = noCol(ActiveCell) ' Calculate the number of columns in the spreadsheet
    result = True
    delCount = 0
    For j = 1 To sizeRow
        For i = 1 To sizeCol
        ' Detect if there is any cell's value is not "-"
            If ActiveCell.Offset(j - delCount - 1, i - 1).Value <> x Then ' We need to adjust the active cells because when one row deleted, the total number of cells changed
                result = False
            End If
        Next i
        If result = True Then
        ' If all of the cells in the row are "-", then delete that row
            ActiveCell.Offset(j - delCount - 1, 0).EntireRow.Delete Shift:=xlUp ' rows adjustment also needed
            delCount = delCount + 1
        End If
        result = True ' Reset the detecter to its original value
    Next j
    
End Sub


Private Function CountHeader() As Integer

' Count how many lines the header breaks into
    Dim counter As Integer
    counter = 1
    Do Until ActiveCell.Offset(counter, 0) = "="
    ' Loop until the "=" we inserted before
        counter = counter + 1
    Loop
    CountHeader = counter ' Return the number of header rows
    
End Function

Private Sub InsertCol()
'
' Insert a Column in the right of the selected column
'
    ActiveCell.EntireColumn.Offset(0, 1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

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
'Count the number of columns in a file
'
    Dim sizeCol As Integer
    sizeCol = Range(c, c.EntireRow.End(xlToRight)).Columns.count
    noCol = sizeCol
End Function

Private Sub DeleteCol(c As Range)
'
' Delete one selected Column
'
    c.EntireColumn.Delete Shift:=xlToLeft
    
End Sub
