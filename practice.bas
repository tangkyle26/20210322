Attribute VB_Name = "Module1"
Sub Demo()
MsgBox "main text", vbYesNoCancel, "我是標題"
End Sub
Sub Demonew()
MsgBox "maintext", 3, ""
End Sub
Sub StringDemo()
Dim i As String
i = "葉庭杏說:大家好，大家好嗎~~~~"
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
Dim userString As String '宣告變數-宣告一個文字型態變數，名稱叫userString
'(a)使用InputBox-InputBox為使用者輸入框的彈跳視窗，第一個引數為視窗主文字
'(b)變數值為輸入框的內容，變數=輸入框函數，即可將輸入框的內容存到變數
userString = InputBox("what is your telephone number?")
'彈跳視窗呈現，字串相加用"&"符號，本行程式為="answer:"和"what is your telephone number"相加
MsgBox "answer:" & userString
End Sub

Sub IntDemo2()
Dim i As Integer
i = 1000
MsgBox i
End Sub

