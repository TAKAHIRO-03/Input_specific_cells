Attribute VB_Name = "Module4"
Sub move_file()


Dim fso As FileSystemObject
Set fso = New FileSystemObject

Call fso.MoveFile("パスを指定")

Set fso = Nothing

End Sub

