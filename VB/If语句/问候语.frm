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
    
    Dim h As Integer
    
    Show
    
    h = Hour(Time)  '取系统小时数
    
    FontSize = 30
    
    ForeColor = RGB(255, 0, 0)
    
    BackColor = RGB(255, 255, 0)
    
    If h < 12 Then
        
        Print "早上好！"
    
    ElseIf h < 18 Then
        
        Print "下午好！"
    
    Else
        
        Print "晚上好！"
    
    End If
    
End Sub
