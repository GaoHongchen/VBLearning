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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ȫ��"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ɾ��"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
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
      Caption         =   "ѡ�޿γ�����"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ѡ�޿γ�"
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
Private Sub Command1_Click()    '���
    
    If Len(Combo1.Text) > 0 Then    '�ж��Ƿ�������
        Combo1.AddItem Combo1.Text
        Text1.Text = Combo1.ListCount
    End If
    
    Combo1.Text = ""
    
    Combo1.SetFocus     '���ý���
       
End Sub

Private Sub Command2_Click()    'ɾ��
    
    Dim ind As Integer
    ind = Combo1.ListIndex
    If ind <> -1 Then       '-1��ʾ��ѡ������
        Combo1.RemoveItem ind   'ɾ����ѡ���ı���
        Text1.Text = Combo1.ListCount
    End If

End Sub

Private Sub Command3_Click()    'ȫ��
    
    Combo1.Clear
    Text1.Text = Combo1.ListCount
    

End Sub

Private Sub Command4_Click()    '�˳�
    
    End
    
End Sub

Private Sub Form_Load()
    
    Combo1.AddItem "��������"
    Combo1.AddItem "��ҳ����"
    Combo1.AddItem "Internet�����̳�"
    Combo1.AddItem "������������"
    Combo1.AddItem "��ý�弼��"
    
    Combo1.Text = ""    '�ÿ���
    
    Text1.Text = Combo1.ListCount   '�������
    
End Sub
