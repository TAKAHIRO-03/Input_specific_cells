VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "キャップ径"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If Me.ListBox1.ListIndex = -1 Then
MsgBox "選択してください"
Else

anss = ListBox1.Value

Unload Me
End If
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
  Worksheets("マスタ").Select
  
  ListBox1.RowSource = Worksheets("マスタ").Range("B2:B3").Address

  Worksheets("【4001】包装資材チェックシ－ト").Select
End Sub

