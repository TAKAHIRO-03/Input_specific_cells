Attribute VB_Name = "Module1"
Sub gattai()

'author: ���ї��O


'============================================================================================================================

'���V�[�g�ǉ�������

'�V�[�g�mmerge�n���폜
    On Error Resume Next
    Application.DisplayAlerts = False
       Worksheets("CSV").delete
    Application.DisplayAlerts = True
    
    
'�V�[�g�mmerge�n����ԉE�ɒǉ�
    Worksheets.Add(after:=Worksheets(Worksheets.count)).Name = "CSV"
'�����������t�@�C��������t�H���_�̏ꏊ c�h���C�u�Ȃ� "C:\test\"
    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
        Range("z2").Value = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End If
    
     Application.ScreenUpdating = False  '��ʂ̕`�ʂ�}������i�}�N���̎��s���x�𑁂�����̂��ړI�j'
        Application.EnableEvents = False  '�C�x���g�̔�����}������'
        
'============================================================================================================================
   

'============================================================================================================================

'���t�@�C���ǂݍ��ݏ�����

Dim Fol As String
Fol = Range("z2").Value & "\"
Dim Fn
Dim NewFile As Workbook
Dim Wb As Workbook
Dim Ws1 As Worksheet
Dim Ws2 As Worksheet
Dim R As Range
Set R = Worksheets("CSV").Range("A1")  '��������`�F�b�N�V�[�g�ɂ���f
Fn = Dir(Fol, vbNormal)
Do Until Fn = ""
Set Wb = Workbooks.Open(Fol & Fn)
'���[�N�V�[�g1���R�s�[����ꍇ�� Wb.Worksheets(1)
'���[�N�V�[�g2���R�s�[����ꍇ�� Wb.Worksheets(2)
Set Ws2 = Wb.Worksheets(1)
'A��1�s�ڂ���8��ڂ܂ł��R�s�[���Č�������
Ws2.Range("A1", Ws2.Cells(Rows.count, 1).End(xlUp)).Resize(, 8).Copy R
Set R = R.End(xlDown).Offset(1)
Wb.Close
'Debug.Print Fn
Fn = Dir
Loop
Set R = Nothing
Set Ws1 = Nothing: Set Ws2 = Nothing
Set Wb = Nothing: Set NewFile = Nothing


 Application.ScreenUpdating = True  '��ʂ̕`�ʂ�}������i�}�N���̎��s���x�𑁂�����̂��ړI�j'
        Application.EnableEvents = True  '�C�x���g�̔�����}������'
End Sub

'============================================================================================================================
