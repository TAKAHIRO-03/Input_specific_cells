VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "�e��"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If Me.ListBox1.ListIndex = -1 Then
MsgBox "�I�����Ă�������"
Else

ans = ListBox1.Value

Unload Me
End If
End Sub


Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
  Worksheets("�}�X�^").Select
  
  ListBox1.RowSource = Worksheets("�}�X�^").Range("C2:C3").Address  '�Q�Ɖӏ�'

  Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Select
End Sub


