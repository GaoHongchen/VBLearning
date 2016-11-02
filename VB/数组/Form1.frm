VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   7755
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   3960
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Private Sub Form_Load()
    
    Dim a As Variant, b(3, 4) As Integer
    
    Dim i As Integer, j As Integer
    
    a = Array(11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22)
    
    For i = 1 To 3
        For j = 1 To 4
            b(i, j) = a((i - 1) * 4 + j) '用一维数组a()的数据存入二维数组b(,)
        Next j
    Next i
    
    Show
    
    For i = 1 To 3      '以矩阵的形式输出到二维数组b(,)
        For j = 1 To 4
            Print b(i, j);    '
            Text1.Text = Text1.Text & Str(b(i, j))
        Next j
        Print
        Text1.Text = Text1.Text & vbCrLf    '加入换行控制符vbCrLf(或Chr(13)+Chr(10))
    Next i
    
End Sub
