VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�������"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
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
      Caption         =   "���룺"
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
    
        MsgBox "��ӭ���û���"
        
    Else
    
        p = MsgBox("�������", 5 + 48, "��������"): Rem ����Ϣ������ʾ�����ԡ��͡�ȡ������ť���Լ�������ͼ��
        
        If p = 4 Then   '4��ʾ�����ˡ����ԡ���ť
            Text1.SetFocus  '���㶨λ��ԭ������ı�����
        Else
            MsgBox "������󣬲������ˣ�"
            End
        End If
        
    End If
        
End Sub

Private Sub Form_Load()

    Text1.PasswordChar = "*"
    Text1.Text = ""
    
End Sub
