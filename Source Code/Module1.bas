Attribute VB_Name = "Module1"
Option Explicit

Sub FilterAndDeleteEmp(rngEmp As Range)
Dim StrSourceSheet As String
Dim lngLastRow As Long
Dim rngFilter As Range
Dim ws As Worksheet

Set ws = ThisWorkbook.Worksheets(rngEmp.Parent.Name)
StrSourceSheet = rngEmp.Parent.Name

Application.ScreenUpdating = False

lngLastRow = ws.Cells(Rows.Count, rngEmp.Column).End(xlUp).Row

If ws.FilterMode Then
    ws.ShowAllData
End If

Set rngFilter = ws.UsedRange

rngFilter.AutoFilter Field:=rngEmp.Column, Criteria1:=rngEmp
rngFilter.SpecialCells(xlCellTypeVisible).EntireRow.Delete

'ws.ShowAllData
rngFilter.AutoFilter

Application.ScreenUpdating = True
End Sub
