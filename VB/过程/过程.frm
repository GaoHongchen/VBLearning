VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7650
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Mysub2()
    Print "*"; Tab(30); "*"
End Sub
Private Sub Mysub1(n)
    Print String(n, "*")
End Sub


Private Sub Form_Load()
    
    Show
    Call Mysub1(30)
    Call Mysub2
    Call Mysub2
    Call Mysub2
    Mysub1 30
    
End Sub
