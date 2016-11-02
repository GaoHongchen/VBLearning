VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim x As String, p As Integer, a As String, b As String
    
    Show
    
    x = InputBox("输入“邮件地址”的内容")
    
    p = InStr(x, "@")   '查找字符@，得到@的位置
    
    a = Left(x, p - 1)  '取@左边部分
    
    b = Mid(x, p + 1)   '取@右边部分，也可用 b=Right(x,Len(x)-p)
    
    Print "用户名：" & a
     
    Print "主机域名：" & b
    
End Sub
