VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "密码检验"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "密码："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Dim p As Integer
    
    If Text1.Text = "123456" Then
    
        MsgBox "欢迎您用机！"
        
    Else
    
        p = MsgBox("密码错误！", 5 + 48, "输入密码"): Rem 在消息框上显示“重试”和“取消”按钮，以及“！”图标
        
        If p = 4 Then   '4表示单击了“重试”按钮
            Text1.SetFocus  '焦点定位在原输入的文本框中
        Else
            MsgBox "密码错误，不重试了！"
            End
        End If
        
    End If
        
End Sub

Private Sub Form_Load()

    Text1.PasswordChar = "*"
    Text1.Text = ""
    
End Sub
