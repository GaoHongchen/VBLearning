VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim x As String, p As Integer, a As String, b As String
    
    Show
    
    x = InputBox("���롰�ʼ���ַ��������")
    
    p = InStr(x, "@")   '�����ַ�@���õ�@��λ��
    
    a = Left(x, p - 1)  'ȡ@��߲���
    
    b = Mid(x, p + 1)   'ȡ@�ұ߲��֣�Ҳ���� b=Right(x,Len(x)-p)
    
    Print "�û�����" & a
     
    Print "����������" & b
    
End Sub
