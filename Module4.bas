Attribute VB_Name = "Module4"
Sub move_file()


Dim fso As FileSystemObject
Set fso = New FileSystemObject

Call fso.MoveFile("C:\Users\riken\Desktop\チェックシート用\*", "C:\Users\riken\Desktop\包装資材フォルダ（past）")

Set fso = Nothing

End Sub

