VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6540
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6540
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ConnectCmd 
      Caption         =   "连接MCGS"
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Pressure 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Temperature 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "压力："
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "温度："
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConnectCmd_Click()

    Pressure.LinkMode = 0
    Temperature.LinkMode = 0
    
    Pressure.LinkTopic = "MCGSRun|DataCentre"
    Temperature.LinkTopic = "MCGSRun|DataCentre"
    
    Pressure.LinkItem = "压力"
    Temperature.LinkItem = "温度"
        
    Pressure.LinkMode = 1
    Temperature.LinkMode = 1
     
    Pressure.Text = 压力
    Temperature.Text = 温度
    
End Sub

