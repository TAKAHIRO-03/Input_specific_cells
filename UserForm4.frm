VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 
'========================================��������ŏ��̒[������============================================
 
Dim Balk As String
Balk = UserForm4.TextBox1.Value

Dim sheet As String
sheet = UserForm4.TextBox3.Value

Dim film As String
film = UserForm4.TextBox5.Value

Dim gap As String
gap = UserForm4.TextBox7.Value

Dim cap As String
cap = UserForm4.TextBox9.Value

Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 12) = Balk
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 27) = cap
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 42) = gap
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 83) = sheet
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(36, 83) = film

'========================================�����܂ōŏ��̒[������============================================


'========================================��������Ō�̒[������============================================

'�o���N
Dim Balks As Range
Set Balks = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 12)
Dim a As Range
Set a = Balks.End(xlDown) ' �A�N�e�B�u�Z���̉��[�̃Z�����擾
Dim valk As String
valk = UserForm4.TextBox2.Value
a = valk

'P�P�[�X
Dim sheets As Range
Set sheets = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 83)
Dim d As Range
Set d = sheets.End(xlDown) ' �A�N�e�B�u�Z���̉��[�̃Z�����擾
Dim shets As String
shets = UserForm4.TextBox4.Value
d = shets

'�V�������N
Dim films As Range
Set films = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(36, 83)
Dim e As Range
Set e = films.End(xlDown) ' �A�N�e�B�u�Z���̉��[�̃Z�����擾
Dim filmss As String
filmss = UserForm4.TextBox6.Value
e = filmss

'�O��
Dim gaps As Range
Set gaps = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 42)
Dim gaps2 As Range
Set gaps2 = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 55)
Dim gaps3 As Range
Set gaps3 = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 68)
Dim last As String
last = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(53, 42).Value
Dim last2 As String
last2 = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 55).Value
Dim last3 As String
last3 = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 68).Value
Dim c As Range
Dim gaap As String
If last = "" And last2 = "" Then
Set c = gaps.End(xlDown) ' �A�N�e�B�u�Z���̉��[�̃Z�����擾
gaap = UserForm4.TextBox8.Value
c = gaap
ElseIf Not last2 = "" And last3 = "" Then
Set c = gaps2.End(xlDown) ' �A�N�e�B�u�Z���̉��[�̃Z�����擾
gaap = UserForm4.TextBox8.Value
c = gaap
ElseIf Not last3 = "" Then
Set c = gaps3.End(xlDown) ' �A�N�e�B�u�Z���̉��[�̃Z�����擾
gaap = UserForm4.TextBox8.Value
c = gaap
End If


'����
Dim caps As Range
Set caps = Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12, 27)
Dim b As Range
Set b = caps.End(xlDown) ' �A�N�e�B�u�Z���̉��[�̃Z�����擾
Dim chap As String
chap = UserForm4.TextBox10.Value
b = chap

'========================================��������Ō�̒[������============================================

'�I��
Unload UserForm4


End Sub

