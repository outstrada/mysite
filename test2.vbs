' Excel�N��
Set oXlsApp = CreateObject("Excel.Application")
If oXlsApp Is Nothing Then
 ' Excel�N�����s
 MsgBox "Excel�N�����s"
Else
 ' Excel�N������
 ' --Excel�\���ifalse�ɂ���Ɣ�\���ɂł���j
 oXlsApp.Application.Visible = true
 ' --Excel�̌x�����\���ɂ���
 oXlsApp.Application.DisplayAlerts = False
 ' --3�b�҂�
 ' WScript.Sleep(3000)
 ' --�u�b�N�ǉ�
 ' oXlsApp.Application.Workbooks.Add()
 ' �u�b�N���J��
 oXlsApp.Application.Workbooks.Open("C:\Users\TOMOHIRO\Desktop\�v���O���~���O\vb\Test.xlsm")
 ' --�V�[�g�I��
 Set oSheet = oXlsApp.Worksheets(1)
 ' --A1�̃Z���ɒl��ݒ�
 ' oSheet.Range("A1").value = "aaa"
 ' ' --�s��2�A��3�̃Z���ɒl��ݒ�
 ' oSheet.Cells(2, 3).value = 100
' �Z�����A�N�e�B�u
 oSheet.Range("D150").Activate
 ' --3�b�҂�
 ' WScript.Sleep(3000)
 ' --Excel�I��
 ' oXlsApp.Quit
 ' --Excel�I�u�W�F�N�g�N���A
 Set oXlsApp = Nothing
End If
