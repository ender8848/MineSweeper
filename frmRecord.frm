VERSION 5.00
Begin VB.Form frmRecord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "扫雷英雄榜"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3705
   Icon            =   "frmRecord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3705
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdexit 
      Caption         =   "关闭"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblname3 
      Caption         =   "初级"
      Height          =   255
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Hard"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lbltime3 
      Caption         =   "初级"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblname2 
      Caption         =   "初级"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "Time"
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Playername"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Easy"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lbltime1 
      Caption         =   "初级"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblname1 
      Caption         =   "初级"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Medium"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lbltime2 
      Caption         =   "初级"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Level"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
    lblname1.Caption = username(0)
    lblname2.Caption = username(1)
    lblname3.Caption = username(2)
    lbltime1.Caption = userscore(0)
    lbltime2.Caption = userscore(1)
    lbltime3.Caption = userscore(2)
    
End Sub

