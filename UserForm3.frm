VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "�i��"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If Me.ListBox1.ListIndex = -1 Then
MsgBox "�I�����Ă�������"
Else

hinsyu = ListBox1.Value

Unload Me

End If

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
  Dim i As Long
  Dim cel As Range
  Dim sht As Worksheet
  Dim last As Long
  Set sht = Worksheets("�}�X�^")
  Set cel = sht.Cells(2, 1)
  last = Worksheets("�}�X�^").Cells(Rows.count, 1).End(xlUp).Row
    
  With ListBox1
   For i = 1 To last '1�`�Ō�̃Z���܂�'
      .AddItem cel
      Set cel = sht.Cells(2 + i, 1)
   Next i
  End With

  Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Select
  
End Sub


Rem      ListBox1.RowSource = Worksheets("�}�X�^").Range("A4:A10").Address

