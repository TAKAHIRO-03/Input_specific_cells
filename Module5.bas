Attribute VB_Name = "Module5"

Rem 宣言エリアで変数を宣言
Rem　パブリック変数を使用

Public anss As String
Public ans As String
Public hinsyu As String
Public keisu As String

Sub Writeinchecksheets()

Rem　ここは触らない　加藤

Call gattai
Call kiridashi
Call delete


Rem　ここまで

Rem　Call Module2.cap　'←inputboxはこちら、動かないように寝かしてます
UserForm1.Show  '←ユーザーフォーム起動用

Rem Call Module2.net　'←inputboxはこちら、動かないように寝かしてます
UserForm2.Show '←ユーザーフォーム起動用

Rem Call Module2.kind　'←inputboxはこちら、動かないように寝かしてます
UserForm3.Show '←ユーザーフォーム起動用

Call Module2.product  '←生産ケース数のinputboxを呼び出します

Dim i As Integer 'For用'
Dim ii As Integer  'バルクセルの値用'
Dim iii As Integer 'シュリンクセルの値用'
Dim iiii As Integer '外栓セルの値用①'
Dim iiiii As Integer '外栓セルの値用②'
Dim iiiiii As Integer '中栓セルの値用'
Dim iiiiiii As Integer '外栓セルの値用③'
Dim iiiiiiii As Integer 'Pケースセルの値用'

Dim mojigirib As String 'バルク①'
Dim mojigiribb As String 'バルク②'
Dim mojigirib_qu As String 'バルク数量'
Dim mojigiris As String 'シュリンク①'
Dim mojigiriss As String 'シュリンク②'
Dim mojigiris_qu As String 'シュリンク数量'
Dim mojigirig As String '外栓①'
Dim mojigiriga As String '外栓②'
Dim mojigiriga_qu As String '外栓数量①'
Dim mojigirigg As String '外栓③'
Dim mojigirigga As String '外栓④'
Dim mojigirigga_qu As String '外栓数量②'
Dim mojigiriggg As String '外栓⑤'
Dim mojigiriggga As String '外栓⑥'
Dim mojigiriggga_qu As String '外栓数量③'
Dim mojigirin As String '中栓①'
Dim mojigirinn As String '中栓②'
Dim mojigirin_qu As String '中栓数量③'
Dim mojigirip As String 'Pケース①'
Dim mojigiripp As String 'Pケース②'
Dim mojigirip_qu As String 'Pケース数量③'

'Cells(縦の値,横の値)'

Application.ScreenUpdating = False

Worksheets("【4001】包装資材チェックシ－ト").Cells(2, 39) = ans
Worksheets("【4001】包装資材チェックシ－ト").Cells(2, 45) = anss
Worksheets("【4001】包装資材チェックシ－ト").Cells(47, 60) = keisu
Worksheets("【4001】包装資材チェックシ－ト").Cells(2, 33) = hinsyu
Worksheets("【4001】包装資材チェックシ－ト").Cells(7, 12) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(8, 12) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(7, 27) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(8, 27) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(7, 42) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(8, 42) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(7, 83) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(8, 83) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(30, 83) = "レ"
Worksheets("【4001】包装資材チェックシ－ト").Cells(31, 83) = "レ"


For i = 1 To Worksheets("CSV").Cells(Rows.count, 1).End(xlUp).Row '1～最後のセルまで'

If Worksheets("CSV").Cells(i, 4).Value = "  松戸工" And ii < 40 Then  'セルを参照して、バルクだったら下記の動作'
    mojigirib_qu = Mid(Worksheets("CSV").Cells(i, 3), 72, 4) '数量を切り取り'
    Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + ii, 12) = mojigirib_qu  '数量代入'
    mojigirib = Mid(Worksheets("CSV").Cells(i, 3), 91, 8) 'QRで読み取った情報を欲しいところだけ、切り出す'
    mojigiribb = Mid(Worksheets("CSV").Cells(i, 3), 115, 2)
    Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + ii, 2) = mojigirib & " - " & mojigiribb '切り取った情報を代入する'
    Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + ii, 9) = "レ"   '150mlと代入'
    ii = ii + 1  '次のセルへ順番に入れるために1増やす'
    
       End If
    
If Worksheets("CSV").Cells(i, 4).Value = "  筑波工" And iii < 16 Then   'シュリンク'
       mojigiris_qu = Mid(Worksheets("CSV").Cells(i, 3), 72, 4) '数量を切り取り'
       Worksheets("【4001】包装資材チェックシ－ト").Cells(36 + iii, 83) = mojigiris_qu  '数量代入'
       mojigiris = Mid(Worksheets("CSV").Cells(i, 3), 105, 7)
       mojigiriss = Mid(Worksheets("CSV").Cells(i, 3), 112, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(36 + iii, 75) = mojigiris & " - " & mojigiriss
       Worksheets("【4001】包装資材チェックシ－ト").Cells(36 + iii, 73) = "レ"
       iii = iii + 1
 
       End If

If Worksheets("CSV").Cells(i, 4).Value = " ＲＶＳオ" And iiii < 42 And iiii < 42 Then   '外栓①'
       mojigiriga_qu = "1200"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiii, 42) = mojigiriga_qu  '数量代入'
       mojigirig = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigiriga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiii, 32) = mojigirig & " -" & mojigiriga
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiii, 39) = "レ"
       iiii = iiii + 1

       End If
       
If Worksheets("CSV").Cells(i, 4).Value = "ＲＶＳオー" And iiii < 42 And iiii < 42 Then   '外栓①'
       mojigiriga_qu = "1200"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiii, 42) = mojigiriga_qu  '数量代入'
       mojigirig = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigiriga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiii, 32) = mojigirig & " -" & mojigiriga
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiii, 39) = "レ"
       iiii = iiii + 1

       End If
    
If Worksheets("CSV").Cells(i, 4).Value = " ＲＶＳオ" And iiii = 42 And iiiii < 42 And iiii = 42 And iiiii < 42 Then   '外栓②'
       mojigirigga_qu = "1200"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 55) = mojigirigga_qu  '数量代入'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 46) = mojigirigg & " -" & mojigirigga
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 52) = "レ"
       iiiii = iiiii + 1

       End If


If Worksheets("CSV").Cells(i, 4).Value = "ＲＶＳオー" And iiii = 42 And iiiii < 42 And iiii = 42 And iiiii < 42 Then   '外栓②'
       mojigirigga_qu = "1200"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 55) = mojigirigga_qu  '数量代入'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 46) = mojigirigg & " -" & mojigirigga
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 52) = "レ"
       iiiii = iiiii + 1

       End If
       
       
If Worksheets("CSV").Cells(i, 4).Value = " ＲＶＳオ" And iiiii = 42 And iiiii = 42 And iiiiiii < 16 And iiiii = 42 And iiiii = 42 And iiiiiii < 16 Then '外栓③'
       mojigiriggga_qu = "1200"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 68) = mojigiriggga_qu  '数量代入'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 60) = mojigirigg & " -" & mojigirigga
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 65) = "レ"
       iiiii = iiiii + 1

       End If


If Worksheets("CSV").Cells(i, 4).Value = "ＲＶＳオー" And iiiii = 42 And iiiii = 42 And iiiiiii < 16 And iiiii = 42 And iiiii = 42 And iiiiiii < 16 Then   '外栓③'
       mojigiriggga_qu = "1200"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 68) = mojigiriggga_qu  '数量代入'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 60) = mojigirigg & " -" & mojigirigga
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiii, 65) = "レ"
       iiiii = iiiii + 1

       End If
       
       
If Worksheets("CSV").Cells(i, 4).Value = " ＲＶＳ中" And iiiiii < 40 Then   '中栓'
       mojigirin_qu = "3000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiii, 27) = mojigirin_qu  '数量代入'
       mojigirin = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirinn = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiii, 17) = mojigirin & " -" & mojigirinn
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiii, 24) = "レ"
       iiiiii = iiiiii + 1
       
             
ElseIf Worksheets("CSV").Cells(i, 4).Value = "ＲＶＳ中栓" And iiiiii < 40 Then   '中栓'
       mojigirin_qu = "3000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiii, 27) = mojigirin_qu  '数量代入'
       mojigirin = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirinn = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiii, 17) = mojigirin & " -" & mojigirinn
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiii, 24) = "レ"
       iiiiii = iiiiii + 1
       
       End If
       
       
  If Worksheets("CSV").Cells(i, 4).Value = "C    " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "160" Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 41, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 59, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "C    " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "159" Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 40, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 58, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "159" Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 44, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 62, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
  
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "161" Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 42, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 60, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "162" Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 42, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 60, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "163" Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 44, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 62, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "C    " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "165" Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 46, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 64, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
   ElseIf Worksheets("CSV").Cells(i, 4).Value = "MC   " And iiiiiiii < 13 Then 'Pケース'
       mojigirip_qu = "1000"
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '数量代入'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 40, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 58, 6)
       Worksheets("【4001】包装資材チェックシ－ト").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
     
              
    End If
    
    Next i
    
Application.ScreenUpdating = True

UserForm4.Show

End Sub

