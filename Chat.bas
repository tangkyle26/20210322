Attribute VB_Name = "Module1"
Sub ChatBot()
Dim userString As String '宣告變數-宣告一個文字型態變數，名稱叫userString
'(a)使用InputBox-InputBox為使用者輸入框的彈跳視窗，第一個引數為視窗主文字
'(b)變數值為輸入框的內容，變數=輸入框函數，即可將輸入框的內容存到變數
userString = InputBox("請問你的名字?")
'彈跳視窗呈現，字串相加用"&"符號，本行程式為="answer:"和"what is your telephone number"相加
MsgBox "Hi " & userString & " 好"
Dim userString1 As String
userString1 = InputBox("你喜歡吃鮭魚嗎?")
MsgBox "我們現在正在優惠喔"
Dim userString2 As String
userString2 = InputBox("那你喜歡吃壽司嗎")
MsgBox "真的嗎~ 那真是太好了!老闆這裡來一盤>>>"
End Sub
