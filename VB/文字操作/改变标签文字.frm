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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command3 
      Caption         =   "��ʾ��ǩ(&D)"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���ر�ǩ(&H)"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ı�������ɫ(&C)"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "����������������"
      BeginProperty Font 
         Name            =   "����"
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
Private Sub Command1_Click()    '�ı�������ɫ

    Clr = Int(15 * Rnd)     '���������ɫ����
    
    Label1.ForeColor = QBColor(Clr)

End Sub

Private Sub Command2_Click()    '���ر�ǩ

    Label1.Visible = False

End Sub

Private Sub Command3_Click()    '��ʾ��ǩ
    
    Label1.Visible = True

End Sub

Private Sub Form_Load()
    
    Randomize   '��ʼ�������������
    
    Label1.BackColor = QBColor(15)  '��ǩ����ɫ
    
    Label1.ForeColor = QBColor(0)   '��ǩǰ��ɫ
    
    Label1.FontSize = 18
    
End Sub
