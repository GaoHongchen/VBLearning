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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label Label1 
      Caption         =   "    ȡ1Ԫ��2Ԫ��5Ԫ��Ӳ�ҹ�10ö������25ԪǮ�����ж����ֲ�ͬ��ȡ����"
      BeginProperty Font 
         Name            =   "����"
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
    
    CurrentX = 0: CurrentY = 1500  'ȷ����ʼ��ʾ��x�����y����
    
    Print , "5Ԫ", "2Ԫ", "1Ԫ"
    
    n = 0   '��¼�������
    
    For a = 0 To 10
        For b = 0 To 10
            c = 10 - b - a
            If a + 2 * b + 5 * c = 25 And c >= 0 Then
                n = n + 1
                Print "("; n; ")", c, b, a
            End If
    Next b, a   '�ϲ�����Next���
    
End Sub
