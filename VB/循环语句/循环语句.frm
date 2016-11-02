VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   6990
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      Caption         =   "    取1元、2元、5元的硬币共10枚，付给25元钱，问有多少种不同的取法？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    
    Show
    
    CurrentX = 0: CurrentY = 1500  '确定开始显示的x坐标和y坐标
    
    Print , "5元", "2元", "1元"
    
    n = 0   '记录解的组数
    
    For a = 0 To 10
        For b = 0 To 10
            c = 10 - b - a
            If a + 2 * b + 5 * c = 25 And c >= 0 Then
                n = n + 1
                Print "("; n; ")", c, b, a
            End If
    Next b, a   '合并两个Next语句
    
End Sub
