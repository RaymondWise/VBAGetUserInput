Attribute VB_Name = "GetUserInput"
Option Explicit
Public Sub test()
    Dim x As Variant
    x = ImportDataFromExternalSource(False, 1)
    Debug.Print x(1)
End Sub

Private Function GetUserInputRange() As Range
    'This is segregated because of how excel handles cancelling a range input
    'See http://stackoverflow.com/a/36630124/1161309
    Dim userAnswer As Range
    On Error GoTo ErrorHandler
    Set userAnswer = Application.InputBox("Please select a range", "Range Selector", Type:=8)
    Set GetUserInputRange = userAnswer
    Exit Function
ErrorHandler:
    Set GetUserInputRange = Nothing
End Function

Private Function ImportDataFromExternalSource(ByVal pickSheet As Boolean, Optional ByVal numberOfColumns As Long = 0) As Variant
    Dim lastRow As Long
    Dim fileName As String
    Dim xlApp As New Application
    Set xlApp = New Excel.Application
    Dim targetBook As Workbook
    Dim targetSheet As Worksheet
    Dim targetDataRange As Range
    On Error GoTo ErrorHandler
    fileName = File_Picker()
    
    Set targetBook = xlApp.Workbooks.Open(fileName)
    Set targetSheet = targetBook.Sheets(1)
    
    If pickSheet Then
        xlApp.ActiveWorkbook.Windows(1).Visible = True
        xlApp.Visible = True
        targetBook.Activate
        targetBook.Sheets(1).Activate
        Set targetDataRange = xlApp.InputBox("Pick a cell on the sheet you would like to import", Type:=8)
        Set targetSheet = targetDataRange.Parent
    End If
    
    If numberOfColumns = 0 Then numberOfColumns = targetSheet.Cells(1, Columns.Count).End(xlToLeft).Column
    lastRow = targetSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ImportDataFromExternalSource = targetSheet.Range(targetSheet.Cells(1, 1), targetSheet.Cells(lastRow, numberOfColumns))
    
CleanExit:
    If pickSheet Then
        xlApp.Quit
        Exit Function
    End If
    ThisWorkbook.Activate
    targetBook.Close
    Exit Function
    
ErrorHandler:
    MsgBox "you've cancelled"
    Resume CleanExit
End Function

Public Function File_Picker() As String
    Dim workbookName As String
    Dim selectFile As FileDialog
    Set selectFile = Application.FileDialog(msoFileDialogOpen)
    With selectFile
        .AllowMultiSelect = False
        .Title = "Select the file with your data."
        .Filters.Clear
        .Filters.Add "Excel Document", ("*.csv, *.xls*")
        .InitialView = msoFileDialogViewDetails
        If .Show Then File_Picker = .SelectedItems(1)
    End With
    Set selectFile = Nothing
End Function




