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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Dim h As Integer
    
    Show
    
    h = Hour(Time)  'ȡϵͳСʱ��
    
    FontSize = 30
    
    ForeColor = RGB(255, 0, 0)
    
    BackColor = RGB(255, 255, 0)
    
    If h < 12 Then
        
        Print "���Ϻã�"
    
    ElseIf h < 18 Then
        
        Print "����ã�"
    
    Else
        
        Print "���Ϻã�"
    
    End If
    
End Sub
