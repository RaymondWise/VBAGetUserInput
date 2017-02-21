Attribute VB_Name = "GetUserInput"
Option Explicit

Public Function GetUserInputRange() As Range
    'This is segregated because of how excel handles cancelling a range input
    'See http://stackoverflow.com/a/36630124/1161309
    On Error GoTo ErrorHandler
    Set GetUserInputRange = Application.InputBox("Please select a range", "Range Selector", Type:=8)
    Exit Function
ErrorHandler:
    Set GetUserInputRange = Nothing
End Function

Public Function ImportDataFromExternalSource(ByVal pickSheet As Boolean, Optional ByVal numberOfColumns As Long = 0) As Variant
    Dim lastRow As Long
    Dim fileName As String
    Dim xlApp As New Excel.Application
    Dim targetBook As Excel.Workbook
    Dim targetSheet As Excel.Worksheet

    Set xlApp = New Excel.Application
    Do While targetBook Is Nothing
        fileName = File_Picker()
        If fileName = "" Then Exit Function
        On Error Resume Next
            Set targetBook = xlApp.Workbooks.Open(fileName)
            If Err <> 0 Then MsgBox "An error occurred while opening the file" & vbNewLine _
                                   & fileName & vbNewLine _
                                   & vbNewLine _
                                   & Err.Description _
                                   , vbCritical
        On Error GoTo 0
    Loop
    Set targetSheet = targetBook.Sheets(1)

    On Error GoTo Cleanfail
        If pickSheet Then
            xlApp.Visible = True
            Set targetSheet = xlApp.InputBox("Pick a cell on the sheet you would like to import", Type:=8).Parent
        End If
    On Error GoTo 0
    
    If numberOfColumns = 0 Then numberOfColumns = targetSheet.Cells(1, targetSheet.Columns.Count).End(xlToLeft).Column
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
    ImportDataFromExternalSource = targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(lastRow, numberOfColumns))

Cleanfail:
    targetBook.Close False
    xlApp.Quit

End Function

Public Function File_Picker() As String
    With Excel.Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Title = "Select the file with your data."
        .Filters.Clear
        .Filters.Add "Excel Document", ("*.csv, *.xls*")
        .InitialView = msoFileDialogViewDetails
        If .Show Then File_Picker = .SelectedItems(1)
    End With
End Function




