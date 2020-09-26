' Author: Gabriel De Jesus
' Purpose: Excel VBA Script that will Calculate and Display Z-Scores, Given an Input of Cells
' Note: This will not work on a single cell, it must be a range of cells... computing a single
'       z-score is incredibly trivial and not worth writing an exception for.
' This code is released under the GNU GPL 3.0 License
' Redistribution must include: Original License, Original Author Name
' Sale is Prohibited without authorization

Option Explicit

Sub FindZScore()
Dim clRange As Range
Dim c As Range, mean As Double, stDev As Double, zScr As Double, clVal As Double, total As Double, dPnts As Integer
Dim rgSource As Range, rgDest As Range
total = 0

Set clRange = Application.InputBox("Select Cells to Compute Z-Scores:", "Compute Z-Scores by G.DeJesus", , , , , , 8)
' Calculate Total Value of Range
For Each c In clRange.Cells
    clVal = c.Value
    total = clVal + total
Next c
dPnts = clRange.Count
' Calculate Mean of Range
mean = total / dPnts
' Calculate Standard Deviation
stDev = Application.WorksheetFunction.stDev(clRange)
' Copy Data To New Sheet
If (Sheet_Exists("Z-Scores") = False) Then
    Sheets.Add.Name = "Z-Scores"
    Else
    MsgBox ("Deleting Previous Z-Score Calculations")
    Sheets("Z-Scores").Delete
    Sheets.Add.Name = "Z-Scores"
End If
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 1).Value = "Data"
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 1).Font.Bold = True
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 2).Value = "Z-Score"
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 2).Font.Bold = True
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 3).Value = "Mean"
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 3).Font.Bold = True
ActiveWorkbook.Sheets("Z-Scores").Cells(2, 3).Value = mean
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 4).Value = "Standard Deviation"
ActiveWorkbook.Sheets("Z-Scores").Cells(1, 4).Font.Bold = True
ActiveWorkbook.Sheets("Z-Scores").Cells(2, 4).Value = stDev
ActiveWorkbook.Sheets("Z-Scores").Columns("A:D").HorizontalAlignment = xlLeft
Set rgSource = clRange
Set rgDest = ActiveWorkbook.Worksheets("Z-Scores").Range("A2")
rgSource.Copy
rgDest.PasteSpecial xlPasteValues
Dim i As Integer
For i = 1 To dPnts
    Cells(i + 1, 2).Value = (Cells(i + 1, 1) - mean) / stDev
Next i
End Sub

Function Sheet_Exists(WorkSheet_Name As String) As Boolean
Dim Work_sheet As Worksheet
Sheet_Exists = False
For Each Work_sheet In ActiveWorkbook.Worksheets
    If Work_sheet.Name = WorkSheet_Name Then
        Sheet_Exists = True
    End If
Next
End Function
