Attribute VB_Name = "Module1"
Sub gattai()

'author: 小林隆弘


'============================================================================================================================

'＜シート追加処理＞

'シート［merge］を削除
    On Error Resume Next
    Application.DisplayAlerts = False
       Worksheets("CSV").delete
    Application.DisplayAlerts = True
    
    
'シート［merge］を一番右に追加
    Worksheets.Add(after:=Worksheets(Worksheets.count)).Name = "CSV"
'結合したいファイルがあるフォルダの場所 cドライブなら "C:\test\"
    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
        Range("z2").Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End If
    
     Application.ScreenUpdating = False  '画面の描写を抑制する（マクロの実行速度を早くするのが目的）'
        Application.EnableEvents = False  'イベントの発生を抑制する'
        
'============================================================================================================================
   

'============================================================================================================================

'＜ファイル読み込み処理＞

Dim Fol As String
Fol = Range("z2").Value & "\"
Dim Fn
Dim NewFile As Workbook
Dim Wb As Workbook
Dim Ws1 As Worksheet
Dim Ws2 As Worksheet
Dim R As Range
Set R = Worksheets("CSV").Range("A1")  '←これをチェックシートにする’
Fn = Dir(Fol, vbNormal)
Do Until Fn = ""
Set Wb = Workbooks.Open(Fol & Fn)
'ワークシート1をコピーする場合は Wb.Worksheets(1)
'ワークシート2をコピーする場合は Wb.Worksheets(2)
Set Ws2 = Wb.Worksheets(1)
'Aの1行目から8列目までをコピーして結合する
Ws2.Range("A1", Ws2.Cells(Rows.count, 1).End(xlUp)).Resize(, 8).Copy R
Set R = R.End(xlDown).Offset(1)
Wb.Close
'Debug.Print Fn
Fn = Dir
Loop
Set R = Nothing
Set Ws1 = Nothing: Set Ws2 = Nothing
Set Wb = Nothing: Set NewFile = Nothing


 Application.ScreenUpdating = True  '画面の描写を抑制する（マクロの実行速度を早くするのが目的）'
        Application.EnableEvents = True  'イベントの発生を抑制する'
End Sub

'============================================================================================================================
