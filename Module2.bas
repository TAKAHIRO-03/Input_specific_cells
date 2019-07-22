Attribute VB_Name = "Module2"
Sub cap()
Rem　ユーザーフォーム表示時は使わない

Do While True
anss = Application.InputBox("10と16どちらですか", "キャップ径確認", Type:=1)
If StrPtr(anss) = 0 Then
   MsgBox "キャンセルします。"
   Exit Do
ElseIf anss = "10" Then
    MsgBox "ありがとうございます。"
    anss = "10"
    Exit Do
ElseIf anss = "16" Then
   MsgBox "ありがとうございます。"
   anss = "16"
   Exit Do
ElseIf Not anss = "" Then
  MsgBox "10か16と入力してください。"
  End If
Loop
End Sub
Sub net()
Rem　ユーザーフォーム表示時は使わない

Do While True
ans = Application.InputBox("150と350どちらですか", "容量確認", Type:=1)
If StrPtr(ans) = 0 Then
   MsgBox "キャンセルします。"
   Exit Do
ElseIf ans = "150" Then
  MsgBox "ありがとうございます。"
    Exit Do
ElseIf ans = "350" Then
   MsgBox "ありがとうございます。"
   Exit Do
ElseIf Not ans = "" Then
  MsgBox "150か350と入力してください。"
  End If
Loop

End Sub
Sub kind()
Rem　ユーザーフォーム表示時は使わない

Do While True
hinsyu = Application.InputBox("今日の品種は？", "品種確認", Type:=2)
If StrPtr(hinsyu) = 0 Then
   MsgBox "キャンセルします。"
   Exit Do
ElseIf hinsyu = "CL477" Then
  MsgBox "ありがとうございます。"
    Exit Do
ElseIf hinsyu = "CL478" Then
  MsgBox "ありがとうございます。"
    Exit Do
ElseIf hinsyu = "CL479" Then
  MsgBox "ありがとうございます。"
    Exit Do
ElseIf hinsyu = "CL480" Then
  MsgBox "ありがとうございます。"
    Exit Do
 ElseIf hinsyu = "CL481" Then
  MsgBox "ありがとうございます。"
    Exit Do
ElseIf hinsyu = "CL482" Then
  MsgBox "ありがとうございます。"
    Exit Do
ElseIf Not ans = "" Then
  MsgBox "正しい品種を入力してください。"
  End If
Loop

End Sub
Sub product()

Do While True
keisu = Application.InputBox("本日のケース数は？", "出来高確認", Type:=1)
If StrPtr(keisu) = 0 Then
   MsgBox "キャンセルします。"
   Exit Do
ElseIf keisu = False Then
   MsgBox "未入力です"
ElseIf keisu = keisu Then
   Rem　MsgBox "ありがとうございます。"
   Exit Do
End If

Loop

End Sub

