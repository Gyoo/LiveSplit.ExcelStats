Attribute VB_Name = "ResetDensity"
Sub ResetDensity()
Dim sht As Worksheet
Dim LastRow As Long
Dim LastColumn As Long

Set sht = ThisWorkbook.Worksheets("Segment History")

sht.Activate

LastRow = sht.UsedRange.Rows(sht.UsedRange.Rows.Count).Row
LastColumn = sht.UsedRange.Columns(sht.UsedRange.Columns.Count).Column

Cells(1, LastColumn + 1) = "Run Finished"

For i = 2 To LastRow
    If (IsEmpty(Cells(i, LastColumn))) Then
        For j = LastColumn To 2 Step -1
            If (IsEmpty(Cells(i, j)) And Not IsEmpty(Cells(i, j - 1))) Then
            Cells(LastRow + 1, j).Value = Cells(LastRow + 1, j).Value + 1
            Exit For
            End If
        Next j
    Else: Cells(LastRow + 1, LastColumn + 1).Value = Cells(LastRow + 1, LastColumn + 1).Value + 1
    End If
Next i

'Graph Generation
Dim ResetChart As Chart

Dim rng As Range
Dim rngTitles As Range
Set rng = Range(Cells(LastRow + 1, 2), Cells(LastRow + 1, LastColumn + 1))

Set ResetChart = Charts.Add
ResetChart.Name = "Reset Density"
For i = 2 To LastColumn + 1
    With ResetChart.SeriesCollection.NewSeries
        .Name = sht.Cells(1, i)
        .Values = sht.Cells(LastRow + 1, i)
    End With
Next i
End Sub
