VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6195
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "ִ��"
      Height          =   615
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "������ɼ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()    'ִ��

    Dim score As Integer, temp As String
    
    score = Val(Text1.Text)
    
    temp = "�ɼ��ȼ�Ϊ��"
    
    Select Case score
        
        Case 0 To 59
            Label2.Caption = temp + "������"
            
        Case 60 To 79
            Label2.Caption = temp + "����"
            
        Case 80 To 100
            Label2.Caption = temp + "����"
            
        Case Else
            Label2.Caption = "�ɼ�����"
            
    End Select
    
End Sub

