VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   7110
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "查找"
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "成绩"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "输入要查找的学号"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Dim xh As Variant, cj As Variant

Private Sub Command1_Click()
    
    Dim key As String, flag As Integer, m As Integer, top As Integer, bott As Integer
    flag = 0
    top = 1: bott = 10
    
    key = Text1.Text    '要查找的学生的学号
    
    Do While top <= bott
        m = Int((top + bott) / 2)
        Select Case True
            Case key = xh(m)
                flag = 1
                Text2.Text = cj(m)
                Exit Do
            Case key < xh(m)
                bott = m - 1
            Case key > xh(m)
                top = m + 1
        End Select
    Loop
    
    If flag = 0 Then
        Text2.Text = "无此学号"
    End If
    
    Text1.SetFocus
    
End Sub

Private Sub Form_Load()
    
    xh = Array("10523", "10623", "11187", "11203", "11205", "11207", "11360", "11402", "11437", "11513")
    
    cj = Array(84, 93, 73, 69, 56, 79, 64, 91, 86, 72)
       
End Sub
