VERSION 5.00
Begin VB.Form frmCustom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Level"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   Icon            =   "frmCustom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtNum 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtCol 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtRow 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "bomb（10-400）"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "col（10-30）"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "rol（10-24）"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim jud As Long
    frmMain.mnuBegin.Checked = False
    frmMain.mnuMiddle.Checked = False
    frmMain.mnuExpert.Checked = False
    frmMain.mnuCust.Checked = True
    
    iCols = Val(txtCol.Text)
    iRows = Val(txtRow.Text)
    iBombs = Val(txtNum.Text)
    jud = Int((iCols * iRows - 81) / 639 * 155 + 25)
    
    
    If iRows < 10 Or iRows > 24 Then
       MsgBox ("行数应在10到24之间")
       Exit Sub
    End If
       
    If iCols < 10 Or iCols > 30 Then
       MsgBox ("列数应在10到30之间")
       Exit Sub
    End If
    
    
    If iBombs < 10 Or iBombs > jud Then
      MsgBox ("雷数应在10到" & jud & "之间的数")
      Exit Sub
    End If
    
    iLevel = 3
    OnGameNew
    frmMain.Form_Paint

    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtRow.Text = 20
    txtCol.Text = 20
    txtNum.Text = 20
End Sub

'改动：txtrow 的背景颜色改成白色
'改动: 修改自定义类型最大雷数的bug
Private Sub Label2_Click()

End Sub
