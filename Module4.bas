Attribute VB_Name = "Module4"
Sub move_file()


Dim fso As FileSystemObject
Set fso = New FileSystemObject

Call fso.MoveFile("C:\Users\riken\Desktop\�`�F�b�N�V�[�g�p\*", "C:\Users\riken\Desktop\����ރt�H���_�ipast�j")

Set fso = Nothing

End Sub

