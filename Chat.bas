Attribute VB_Name = "Module1"
Sub ChatBot()
Dim userString As String '�ŧi�ܼ�-�ŧi�@�Ӥ�r���A�ܼơA�W�٥suserString
'(a)�ϥ�InputBox-InputBox���ϥΪ̿�J�ت��u�������A�Ĥ@�Ӥ޼Ƭ������D��r
'(b)�ܼƭȬ���J�ت����e�A�ܼ�=��J�ب�ơA�Y�i�N��J�ت����e�s���ܼ�
userString = InputBox("�аݧA���W�r?")
'�u�������e�{�A�r��ۥ[��"&"�Ÿ��A����{����="answer:"�M"what is your telephone number"�ۥ[
MsgBox "Hi " & userString & " �n"
Dim userString1 As String
userString1 = InputBox("�A���w�Y�D����?")
MsgBox "�ڭ̲{�b���b�u�f��"
Dim userString2 As String
userString2 = InputBox("���A���w�Y�إq��")
MsgBox "�u����~ ���u�O�Ӧn�F!����o�̨Ӥ@�L>>>"
End Sub
