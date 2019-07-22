VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "UserForm4"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8745
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 
'========================================ここから最初の端数入力============================================
 
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

Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 12) = Balk
Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 27) = cap
Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 42) = gap
Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 83) = sheet
Worksheets("【4001】包装資材チェックシ－ト").Cells(36, 83) = film

'========================================ここまで最初の端数入力============================================


'========================================ここから最後の端数入力============================================

'バルク
Dim Balks As Range
Set Balks = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 12)
Dim a As Range
Set a = Balks.End(xlDown) ' アクティブセルの下端のセルを取得
Dim valk As String
valk = UserForm4.TextBox2.Value
a = valk

'Pケース
Dim sheets As Range
Set sheets = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 83)
Dim d As Range
Set d = sheets.End(xlDown) ' アクティブセルの下端のセルを取得
Dim shets As String
shets = UserForm4.TextBox4.Value
d = shets

'シュリンク
Dim films As Range
Set films = Worksheets("【4001】包装資材チェックシ－ト").Cells(36, 83)
Dim e As Range
Set e = films.End(xlDown) ' アクティブセルの下端のセルを取得
Dim filmss As String
filmss = UserForm4.TextBox6.Value
e = filmss

'外栓
Dim gaps As Range
Set gaps = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 42)
Dim gaps2 As Range
Set gaps2 = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 55)
Dim gaps3 As Range
Set gaps3 = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 68)
Dim last As String
last = Worksheets("【4001】包装資材チェックシ－ト").Cells(53, 42).Value
Dim last2 As String
last2 = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 55).Value
Dim last3 As String
last3 = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 68).Value
Dim c As Range
Dim gaap As String
If last = "" And last2 = "" Then
Set c = gaps.End(xlDown) ' アクティブセルの下端のセルを取得
gaap = UserForm4.TextBox8.Value
c = gaap
ElseIf Not last2 = "" And last3 = "" Then
Set c = gaps2.End(xlDown) ' アクティブセルの下端のセルを取得
gaap = UserForm4.TextBox8.Value
c = gaap
ElseIf Not last3 = "" Then
Set c = gaps3.End(xlDown) ' アクティブセルの下端のセルを取得
gaap = UserForm4.TextBox8.Value
c = gaap
End If


'中栓
Dim caps As Range
Set caps = Worksheets("【4001】包装資材チェックシ－ト").Cells(12, 27)
Dim b As Range
Set b = caps.End(xlDown) ' アクティブセルの下端のセルを取得
Dim chap As String
chap = UserForm4.TextBox10.Value
b = chap

'========================================ここから最後の端数入力============================================

'終了
Unload UserForm4


End Sub

