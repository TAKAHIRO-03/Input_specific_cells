Attribute VB_Name = "Module2"
Sub cap()
Rem�@���[�U�[�t�H�[���\�����͎g��Ȃ�

Do While True
anss = Application.InputBox("10��16�ǂ���ł���", "�L���b�v�a�m�F", Type:=1)
If StrPtr(anss) = 0 Then
   MsgBox "�L�����Z�����܂��B"
   Exit Do
ElseIf anss = "10" Then
    MsgBox "���肪�Ƃ��������܂��B"
    anss = "10"
    Exit Do
ElseIf anss = "16" Then
   MsgBox "���肪�Ƃ��������܂��B"
   anss = "16"
   Exit Do
ElseIf Not anss = "" Then
  MsgBox "10��16�Ɠ��͂��Ă��������B"
  End If
Loop
End Sub
Sub net()
Rem�@���[�U�[�t�H�[���\�����͎g��Ȃ�

Do While True
ans = Application.InputBox("150��350�ǂ���ł���", "�e�ʊm�F", Type:=1)
If StrPtr(ans) = 0 Then
   MsgBox "�L�����Z�����܂��B"
   Exit Do
ElseIf ans = "150" Then
  MsgBox "���肪�Ƃ��������܂��B"
    Exit Do
ElseIf ans = "350" Then
   MsgBox "���肪�Ƃ��������܂��B"
   Exit Do
ElseIf Not ans = "" Then
  MsgBox "150��350�Ɠ��͂��Ă��������B"
  End If
Loop

End Sub
Sub kind()
Rem�@���[�U�[�t�H�[���\�����͎g��Ȃ�

Do While True
hinsyu = Application.InputBox("�����̕i��́H", "�i��m�F", Type:=2)
If StrPtr(hinsyu) = 0 Then
   MsgBox "�L�����Z�����܂��B"
   Exit Do
ElseIf hinsyu = "CL477" Then
  MsgBox "���肪�Ƃ��������܂��B"
    Exit Do
ElseIf hinsyu = "CL478" Then
  MsgBox "���肪�Ƃ��������܂��B"
    Exit Do
ElseIf hinsyu = "CL479" Then
  MsgBox "���肪�Ƃ��������܂��B"
    Exit Do
ElseIf hinsyu = "CL480" Then
  MsgBox "���肪�Ƃ��������܂��B"
    Exit Do
 ElseIf hinsyu = "CL481" Then
  MsgBox "���肪�Ƃ��������܂��B"
    Exit Do
ElseIf hinsyu = "CL482" Then
  MsgBox "���肪�Ƃ��������܂��B"
    Exit Do
ElseIf Not ans = "" Then
  MsgBox "�������i�����͂��Ă��������B"
  End If
Loop

End Sub
Sub product()

Do While True
keisu = Application.InputBox("�{���̃P�[�X���́H", "�o�����m�F", Type:=1)
If StrPtr(keisu) = 0 Then
   MsgBox "�L�����Z�����܂��B"
   Exit Do
ElseIf keisu = False Then
   MsgBox "�����͂ł�"
ElseIf keisu = keisu Then
   Rem�@MsgBox "���肪�Ƃ��������܂��B"
   Exit Do
End If

Loop

End Sub

