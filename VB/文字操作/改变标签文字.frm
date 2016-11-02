VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "显示标签(&D)"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "隐藏标签(&H)"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "改变文字颜色(&C)"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "计算机程序设计语言"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   600
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()    '改变文字颜色

    Clr = Int(15 * Rnd)     '产生随机颜色代码
    
    Label1.ForeColor = QBColor(Clr)

End Sub

Private Sub Command2_Click()    '隐藏标签

    Label1.Visible = False

End Sub

Private Sub Command3_Click()    '显示标签
    
    Label1.Visible = True

End Sub

Private Sub Form_Load()
    
    Randomize   '初始化随机数发生器
    
    Label1.BackColor = QBColor(15)  '标签背景色
    
    Label1.ForeColor = QBColor(0)   '标签前景色
    
    Label1.FontSize = 18
    
End Sub
