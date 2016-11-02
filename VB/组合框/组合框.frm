VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "全清"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "选修课程总数"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "选修课程"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()    '添加
    
    If Len(Combo1.Text) > 0 Then    '判断是否有内容
        Combo1.AddItem Combo1.Text
        Text1.Text = Combo1.ListCount
    End If
    
    Combo1.Text = ""
    
    Combo1.SetFocus     '设置焦点
       
End Sub

Private Sub Command2_Click()    '删除
    
    Dim ind As Integer
    ind = Combo1.ListIndex
    If ind <> -1 Then       '-1表示无选定表项
        Combo1.RemoveItem ind   '删除已选定的表项
        Text1.Text = Combo1.ListCount
    End If

End Sub

Private Sub Command3_Click()    '全清
    
    Combo1.Clear
    Text1.Text = Combo1.ListCount
    

End Sub

Private Sub Command4_Click()    '退出
    
    End
    
End Sub

Private Sub Form_Load()
    
    Combo1.AddItem "电子商务"
    Combo1.AddItem "网页制作"
    Combo1.AddItem "Internet简明教程"
    Combo1.AddItem "计算机网络基础"
    Combo1.AddItem "多媒体技术"
    
    Combo1.Text = ""    '置空置
    
    Text1.Text = Combo1.ListCount   '表项个数
    
End Sub
