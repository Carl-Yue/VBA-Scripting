VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Import with all subs or functions for the purpose of importing data

Sub ImportFiles()
'
' Import files from a folder that can be selected
'
    Dim iFolderPath As String
    Dim iFileDialog As FileDialog
    Dim iFilePath As String
    Dim FilePath() As String
    Dim counter As Integer
    On Error GoTo ErrHandler ' If error happend, then go to ErrHandler
    Set iFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    ' A dialog pops out for folder selection
    iFileDialog.AllowMultiSelect = False ' one folder only
    iFileDialog.Title = "Select a folder: " ' dialog title
    If iFileDialog.Show = -1 Then
    ' If one folder selected, then same the path
        iFolderPath = iFileDialog.SelectedItems(1)
    ElseIf iFileDialog.Show = 0 Then
    ' If no folder selected, then exit the sub
        Exit Sub
    End If
    If iFolderPath = "" Then Exit Sub ' If path is empty, then exit the sub
    Application.ScreenUpdating = False ' Stop Screen update when importing files
    iFilePath = Dir(iFolderPath & "\*.txt") ' load files name into the file path
    Do While iFilePath <> ""
    ' loop until no files left in the folder
        counter = counter + 1
        ActiveWorkbook.Sheets.Add Before:=Worksheets(Worksheets.count)
        If UCase(iFilePath) Like "*ISERIES*" Then
            If UCase(iFilePath) Like "*DEV*" Then
                ActiveSheet.Name = "ISeries DEV"
            ElseIf UCase(iFilePath) Like "*QA*" Then
                ActiveSheet.Name = "ISeries QA"
            ElseIf UCase(iFilePath) Like "*PROD*" Then
                ActiveSheet.Name = "ISeries Prod"
            End If
        ElseIf UCase(iFilePath) Like "*TANDEM*" Then
            If UCase(iFilePath) Like "*PROD*" Then
                ActiveSheet.Name = "Tandem Prod"
            ElseIf UCase(iFilePath) Like "*QA AND DEV*" Then
                ActiveSheet.Name = "Tandem QA and DEV"
            ElseIf UCase(iFilePath) Like "*DEV*" Then
                ActiveSheet.Name = "Tandem DEV"
            End If
        Else
            ActiveSheet.Name = iFilePath
        End If
        
        Sheets(counter).Select ' select the spreadsheets one by one
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" _
          & iFolderPath & "\" & iFilePath, Destination:=Range("A1"))
            ' property definitions: useful when files are in delimited form
            .Name = "a" & count
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = True
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileOtherDelimiter = "|"
            .TextFileColumnDataTypes = Array(1, 1, 1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With
        ' Range("O3").Value = iFilePath ' store the file title in some cell
        Columns("A").Select ' reset selection to the first column
        If UCase(iFilePath) Like "*TANDEM*" Then
        ' If the file is Tandem file, then split columns as Tandem files
            SplitTFile
        ElseIf UCase(iFilePath) Like "*ISERIES*" Then
        ' If the file is iSeries file, then split columns as iSeries files
            SplitIFile
        ' More files types can be put here
        End If
                If (UCase(iFilePath) Like "*TANDEM*") Or (UCase(iFilePath) Like "*ISERIES*") Then
            Rows("1").Delete Shift:=xlUp
            Rows("1").Delete Shift:=xlUp
        End If
        iFilePath = Dir
        Application.ScreenUpdating = True
    Loop
    
    Exit Sub
ErrHandler:
    MsgBox "no files txt", , "Result: "
End Sub

Private Sub SplitTFile()
'
' Split Tandam-file Columns
'
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(21, 1), Array(36, 1), Array(57, 1), Array(68, 1), _
        Array(79, 1), Array(89, 1), Array(104, 1), Array(121, 1), Array(138, 1), Array(155, 1), _
        Array(173, 1)), TrailingMinusNumbers:=True
End Sub

Private Sub SplitIFile()
'
' Split iSeries-file Columns
'
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(23, 1), Array(47, 1), Array(59, 1), Array(67, 1), _
        Array(81, 1), Array(92, 1), Array(115, 1), Array(141, 1)), TrailingMinusNumbers:= _
        True
End Sub

Sub SplitWorkbook()

Dim FileExtStr As String
Dim FileFormatNum As Long
Dim xWs As Worksheet
Dim xWb As Workbook
Dim FolderName As String
Application.ScreenUpdating = False
Set xWb = Application.ThisWorkbook
DateString = Format(Now, "yyyy-mm-dd hh-mm-ss")
FolderName = xWb.Path & "\" & xWb.Name & " " & DateString
MkDir FolderName
For Each xWs In xWb.Worksheets
    xWs.Copy
    If Val(Application.Version) < 12 Then
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        Select Case xWb.FileFormat
            Case 51:
                FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If Application.ActiveWorkbook.HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56:
                FileExtStr = ".xls": FileFormatNum = 56
            Case Else:
                FileExtStr = ".xlsb": FileFormatNum = 50
        End Select
    End If
    xFile = FolderName & "\" & Application.ActiveWorkbook.Sheets(1).Name & FileExtStr
    Application.ActiveWorkbook.SaveAs xFile, FileFormat:=FileFormatNum
    Application.ActiveWorkbook.Close False
Next
Application.ScreenUpdating = True

End Sub








