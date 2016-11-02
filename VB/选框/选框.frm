VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   6660
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6120
      Top             =   120
   End
   Begin VB.CheckBox Check4 
      Caption         =   "红色"
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "斜体"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "25号字"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "黑体"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   975
   End
   Begin VB.OptionButton Option3 
      Caption         =   "楷体"
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "幼圆"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "宋体"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label TimerLabel 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "请在文本框中输入文字："
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

    If Check1.Value = 1 Then      '判断复选框是否被选中
        Text1.FontName = "黑体"
    Else
        Text1.FontName = "宋体"
    End If
    
End Sub

Private Sub Check3_Click()

    If Check3.Value = 1 Then
        Text1.FontItalic = True
    Else
        Text1.FontItalic = False
    End If
        
End Sub

Private Sub Check2_Click()

    If Check2.Value = 1 Then
        Text1.FontSize = 25
    Else
        Text1.FontSize = 9
    End If
        
End Sub

Private Sub Check4_Click()

    If Check4.Value = 1 Then
        Text1.ForeColor = RGB(255, 0, 0)
    Else
        Text1.ForeColor = RGB(0, 0, 0)
    End If
        
End Sub

Private Sub Option1_Click()

    Text1.FontName = "宋体"

End Sub

Private Sub Option2_Click()

    Text1.FontName = "幼圆"

End Sub

Private Sub Option3_Click()

    Text1.FontName = "楷体_GB2312"

End Sub

Private Sub Timer1_Timer()

    TimerLabel.Caption = Time
    
End Sub
