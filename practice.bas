Attribute VB_Name = "Module1"
Sub Demo()
MsgBox "main text", vbYesNoCancel, "�ڬO���D"
End Sub
Sub Demonew()
MsgBox "maintext", 3, ""
End Sub
Sub StringDemo()
Dim i As String
i = "���x����:�j�a�n�A�j�a�n��~~~~"
MsgBox i
End Sub
Sub SingleDemo()
Dim i As Single
i = 120.999
MsgBox i
End Sub
Sub DoubleDemo()
Dim i As Double
i = 120.999651651365
MsgBox i
End Sub
Sub DateDemo()
Dim a As Date
a = Now
MsgBox a
End Sub
Sub BooleanDemo()
Dim a As Boolean
a = True
MsgBox a
End Sub
Sub inputBoxDemo()
Dim userString As String '�ŧi�ܼ�-�ŧi�@�Ӥ�r���A�ܼơA�W�٥suserString
'(a)�ϥ�InputBox-InputBox���ϥΪ̿�J�ت��u�������A�Ĥ@�Ӥ޼Ƭ������D��r
'(b)�ܼƭȬ���J�ت����e�A�ܼ�=��J�ب�ơA�Y�i�N��J�ت����e�s���ܼ�
userString = InputBox("what is your telephone number?")
'�u�������e�{�A�r��ۥ[��"&"�Ÿ��A����{����="answer:"�M"what is your telephone number"�ۥ[
MsgBox "answer:" & userString
End Sub

Sub IntDemo2()
Dim i As Integer
i = 1000
MsgBox i
End Sub

