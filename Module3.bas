Attribute VB_Name = "Module3"
Sub kiridashi()

Dim kirikiri(1 To 300)
Dim count(1 To 300)
Dim mojigirib As String
Dim mojigiribb As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

Application.ScreenUpdating = False

Worksheets("CSV").Activate

For i = 1 To Worksheets("CSV").Cells(Rows.count, 1).End(xlUp).Row

kirikiri(i) = Mid(Worksheets("CSV").Cells(i, 3), 27, 5)
Worksheets("CSV").Cells(i, 4) = kirikiri(i)
count(i) = Len(Worksheets("CSV").Cells(i, 3))
Worksheets("CSV").Cells(i, 5) = count(i)

Next i

For j = 1 To Worksheets("CSV").Cells(Rows.count, 1).End(xlUp).Row - 1

For k = Worksheets("CSV").Cells(Rows.count, 1).End(xlUp).Row To j + 1 Step -1

If Worksheets("CSV").Cells(j, 3).Value = Worksheets("CSV").Cells(k, 3).Value Then

Rows(k).delete
k = k - 1

End If

Next k

Next j

Application.ScreenUpdating = True
    
End Sub




