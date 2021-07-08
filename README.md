# 四則運算
程式(VBA)
Public Sub 四則運算()
Rem 輸入成績
x = InputBox("輸入學生姓名")
y = InputBox("輸入國文分數")
Z = InputBox("輸入微積分分數")
q = InputBox("輸入計概分數")
Rem 計算加權平均
S = (y * 2 + Z * 4 + q * 3) / 9
Rem 顯示成績
MsgBox "學生姓名:" & x & "國文成績:" & y & "微積分分數:" & Z & "計概分數:" & q & "加權平均數:" & S


If S >= 60 Then
MsgBox "及格"
ElseIf S < 50 Then
MsgBox ("不及格")
Else
MsgBox ("補考")
End If
End Sub
