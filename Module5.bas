Attribute VB_Name = "Module5"

Rem �錾�G���A�ŕϐ���錾
Rem�@�p�u���b�N�ϐ����g�p

Public anss As String
Public ans As String
Public hinsyu As String
Public keisu As String

Sub Writeinchecksheets()

Rem�@�����͐G��Ȃ��@����

Call gattai
Call kiridashi
Call delete


Rem�@�����܂�

Rem�@Call Module2.cap�@'��inputbox�͂�����A�����Ȃ��悤�ɐQ�����Ă܂�
UserForm1.Show  '�����[�U�[�t�H�[���N���p

Rem Call Module2.net�@'��inputbox�͂�����A�����Ȃ��悤�ɐQ�����Ă܂�
UserForm2.Show '�����[�U�[�t�H�[���N���p

Rem Call Module2.kind�@'��inputbox�͂�����A�����Ȃ��悤�ɐQ�����Ă܂�
UserForm3.Show '�����[�U�[�t�H�[���N���p

Call Module2.product  '�����Y�P�[�X����inputbox���Ăяo���܂�

Dim i As Integer 'For�p'
Dim ii As Integer  '�o���N�Z���̒l�p'
Dim iii As Integer '�V�������N�Z���̒l�p'
Dim iiii As Integer '�O���Z���̒l�p�@'
Dim iiiii As Integer '�O���Z���̒l�p�A'
Dim iiiiii As Integer '�����Z���̒l�p'
Dim iiiiiii As Integer '�O���Z���̒l�p�B'
Dim iiiiiiii As Integer 'P�P�[�X�Z���̒l�p'

Dim mojigirib As String '�o���N�@'
Dim mojigiribb As String '�o���N�A'
Dim mojigirib_qu As String '�o���N����'
Dim mojigiris As String '�V�������N�@'
Dim mojigiriss As String '�V�������N�A'
Dim mojigiris_qu As String '�V�������N����'
Dim mojigirig As String '�O���@'
Dim mojigiriga As String '�O���A'
Dim mojigiriga_qu As String '�O�𐔗ʇ@'
Dim mojigirigg As String '�O���B'
Dim mojigirigga As String '�O���C'
Dim mojigirigga_qu As String '�O�𐔗ʇA'
Dim mojigiriggg As String '�O���D'
Dim mojigiriggga As String '�O���E'
Dim mojigiriggga_qu As String '�O�𐔗ʇB'
Dim mojigirin As String '�����@'
Dim mojigirinn As String '�����A'
Dim mojigirin_qu As String '���𐔗ʇB'
Dim mojigirip As String 'P�P�[�X�@'
Dim mojigiripp As String 'P�P�[�X�A'
Dim mojigirip_qu As String 'P�P�[�X���ʇB'

'Cells(�c�̒l,���̒l)'

Application.ScreenUpdating = False

Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(2, 39) = ans
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(2, 45) = anss
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(47, 60) = keisu
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(2, 33) = hinsyu
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(7, 12) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(8, 12) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(7, 27) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(8, 27) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(7, 42) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(8, 42) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(7, 83) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(8, 83) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(30, 83) = "��"
Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(31, 83) = "��"


For i = 1 To Worksheets("CSV").Cells(Rows.count, 1).End(xlUp).Row '1�`�Ō�̃Z���܂�'

If Worksheets("CSV").Cells(i, 4).Value = "  ���ˍH" And ii < 40 Then  '�Z�����Q�Ƃ��āA�o���N�������牺�L�̓���'
    mojigirib_qu = Mid(Worksheets("CSV").Cells(i, 3), 72, 4) '���ʂ�؂���'
    Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + ii, 12) = mojigirib_qu  '���ʑ��'
    mojigirib = Mid(Worksheets("CSV").Cells(i, 3), 91, 8) 'QR�œǂݎ��������~�����Ƃ��낾���A�؂�o��'
    mojigiribb = Mid(Worksheets("CSV").Cells(i, 3), 115, 2)
    Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + ii, 2) = mojigirib & " - " & mojigiribb '�؂���������������'
    Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + ii, 9) = "��"   '150ml�Ƒ��'
    ii = ii + 1  '���̃Z���֏��Ԃɓ���邽�߂�1���₷'
    
       End If
    
If Worksheets("CSV").Cells(i, 4).Value = "  �}�g�H" And iii < 16 Then   '�V�������N'
       mojigiris_qu = Mid(Worksheets("CSV").Cells(i, 3), 72, 4) '���ʂ�؂���'
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(36 + iii, 83) = mojigiris_qu  '���ʑ��'
       mojigiris = Mid(Worksheets("CSV").Cells(i, 3), 105, 7)
       mojigiriss = Mid(Worksheets("CSV").Cells(i, 3), 112, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(36 + iii, 75) = mojigiris & " - " & mojigiriss
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(36 + iii, 73) = "��"
       iii = iii + 1
 
       End If

If Worksheets("CSV").Cells(i, 4).Value = " �q�u�r�I" And iiii < 42 And iiii < 42 Then   '�O���@'
       mojigiriga_qu = "1200"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiii, 42) = mojigiriga_qu  '���ʑ��'
       mojigirig = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigiriga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiii, 32) = mojigirig & " -" & mojigiriga
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiii, 39) = "��"
       iiii = iiii + 1

       End If
       
If Worksheets("CSV").Cells(i, 4).Value = "�q�u�r�I�[" And iiii < 42 And iiii < 42 Then   '�O���@'
       mojigiriga_qu = "1200"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiii, 42) = mojigiriga_qu  '���ʑ��'
       mojigirig = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigiriga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiii, 32) = mojigirig & " -" & mojigiriga
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiii, 39) = "��"
       iiii = iiii + 1

       End If
    
If Worksheets("CSV").Cells(i, 4).Value = " �q�u�r�I" And iiii = 42 And iiiii < 42 And iiii = 42 And iiiii < 42 Then   '�O���A'
       mojigirigga_qu = "1200"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 55) = mojigirigga_qu  '���ʑ��'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 46) = mojigirigg & " -" & mojigirigga
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 52) = "��"
       iiiii = iiiii + 1

       End If


If Worksheets("CSV").Cells(i, 4).Value = "�q�u�r�I�[" And iiii = 42 And iiiii < 42 And iiii = 42 And iiiii < 42 Then   '�O���A'
       mojigirigga_qu = "1200"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 55) = mojigirigga_qu  '���ʑ��'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 46) = mojigirigg & " -" & mojigirigga
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 52) = "��"
       iiiii = iiiii + 1

       End If
       
       
If Worksheets("CSV").Cells(i, 4).Value = " �q�u�r�I" And iiiii = 42 And iiiii = 42 And iiiiiii < 16 And iiiii = 42 And iiiii = 42 And iiiiiii < 16 Then '�O���B'
       mojigiriggga_qu = "1200"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 68) = mojigiriggga_qu  '���ʑ��'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 60) = mojigirigg & " -" & mojigirigga
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 65) = "��"
       iiiii = iiiii + 1

       End If


If Worksheets("CSV").Cells(i, 4).Value = "�q�u�r�I�[" And iiiii = 42 And iiiii = 42 And iiiiiii < 16 And iiiii = 42 And iiiii = 42 And iiiiiii < 16 Then   '�O���B'
       mojigiriggga_qu = "1200"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 68) = mojigiriggga_qu  '���ʑ��'
       mojigirigg = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirigga = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 60) = mojigirigg & " -" & mojigirigga
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiii, 65) = "��"
       iiiii = iiiii + 1

       End If
       
       
If Worksheets("CSV").Cells(i, 4).Value = " �q�u�r��" And iiiiii < 40 Then   '����'
       mojigirin_qu = "3000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiii, 27) = mojigirin_qu  '���ʑ��'
       mojigirin = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirinn = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiii, 17) = mojigirin & " -" & mojigirinn
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiii, 24) = "��"
       iiiiii = iiiiii + 1
       
             
ElseIf Worksheets("CSV").Cells(i, 4).Value = "�q�u�r����" And iiiiii < 40 Then   '����'
       mojigirin_qu = "3000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiii, 27) = mojigirin_qu  '���ʑ��'
       mojigirin = Mid(Worksheets("CSV").Cells(i, 3), 10, 8)
       mojigirinn = Mid(Worksheets("CSV").Cells(i, 3), 18, 3)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiii, 17) = mojigirin & " -" & mojigirinn
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiii, 24) = "��"
       iiiiii = iiiiii + 1
       
       End If
       
       
  If Worksheets("CSV").Cells(i, 4).Value = "C    " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "160" Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 41, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 59, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "C    " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "159" Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 40, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 58, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "159" Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 44, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 62, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
  
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "161" Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 42, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 60, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "162" Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 42, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 60, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "     " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "163" Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 44, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 62, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
  ElseIf Worksheets("CSV").Cells(i, 4).Value = "C    " And iiiiiiii < 13 And Worksheets("CSV").Cells(i, 5).Value = "165" Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 46, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 64, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
       
   ElseIf Worksheets("CSV").Cells(i, 4).Value = "MC   " And iiiiiiii < 13 Then 'P�P�[�X'
       mojigirip_qu = "1000"
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 83) = mojigirip_qu  '���ʑ��'
       mojigirip = Mid(Worksheets("CSV").Cells(i, 3), 40, 8)
       mojigiripp = Mid(Worksheets("CSV").Cells(i, 3), 58, 6)
       Worksheets("�y4001�z����ރ`�F�b�N�V�|�g").Cells(12 + iiiiiiii, 73) = mojigirip & " - " & mojigiripp
       iiiiiiii = iiiiiiii + 1
     
              
    End If
    
    Next i
    
Application.ScreenUpdating = True

UserForm4.Show

End Sub

