VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mine"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   0
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game(G)"
      Begin VB.Menu mnuStart 
         Caption         =   "start"
         Shortcut        =   {F2}
      End
      Begin VB.Menu sp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBegin 
         Caption         =   "Easy(&B)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuMiddle 
         Caption         =   "medium"
      End
      Begin VB.Menu mnuExpert 
         Caption         =   "Hard"
      End
      Begin VB.Menu mnuCust 
         Caption         =   "Custom"
      End
      Begin VB.Menu sp3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecord 
         Caption         =   "Leaderboard"
      End
      Begin VB.Menu sp2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Mine"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Integer, hang As Integer, lie As Integer



Private Sub Form_Load()
frmMain.Show
    Dim i As Integer
    
    For i = 0 To 15
        Set picBoard(i) = LoadResPicture(100 + i, 0)           'load "bomb" picture from mine.res
    Next
    
    For i = 0 To 10
        Set picNum(i) = LoadResPicture(120 + i, 0)             'load "number" picture (0-9)
    Next
    
    
    For i = 0 To 4
        Set picFace(i) = LoadResPicture(140 + i, 0)            'load "smile face" picture
    Next
 
    
    ReDim arrBomb(0 To 9, 0 To 9)                           'the difult leve is "easy" level (10arrows, 10lines)
    
    Call OnGameNew
    
    Call Form_Paint
    
    Call ReadRecord
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ln  As Integer, col As Integer
    
    If Button = 1 Then
    
        If x >= ptFace.x And y >= ptFace.y And x <= ptFace.x + 24 And y <= ptFace.y + 24 Then   'if click on face,change the picture to mimic3d special effect
            bLButtonDown = True
            bDrawFace = True
            DrawFace SMILEDOWN
            Exit Sub
        End If
        
        If iFace = CRY Or iFace = COOL Then Exit Sub              'if click out of the face, or game has already over, exit sub
        
        bLButtonDown = True                                       'if click out of the face when game is running, change the expression (quite funny)
        DrawFace SURPRISE
        
        If x < 12 Or x >= 12 + 16 * iCols Or y < 55 Or y >= 55 + 16 * iRows Then Exit Sub    'if click out of board, exit sub
        
        ln = Int((y - 40) / 16) - 1                                    'calculate which board is clicled according the click position
        col = Int((x + 4) / 16) - 1
        
        iOldRow = ln
        iOldCol = col
          
        If bRButtonDown Then                                        'if a bomb has been flaged, store the flag info
          DrawNineBoard col, ln
        End If
     
        If arrBomb(col, ln).board = MARK Then                     '3d vspecial effect, change the picture of question mark if left click on board
            DrawBoard col, ln, MARKDOWN
        End If
        If arrBomb(col, ln).board = COMMON Then                   '3d vspecial effect, change the picture of board if left click on board
            DrawBoard col, ln, DOWN
        End If
    
    ElseIf Button = 2 Then                                     'different mark for right click
        If (iFace = CRY Or iFace = COOL) Then Exit Sub         'if the game has already been gameover, exit sub
        If x < 12 Or x >= 12 + 16 * iCols Or y < 55 Or y >= 55 + 16 * iRows Then Exit Sub
                                     ' 在board外面点击不响应
        bRButtonDown = True
        
        ln = Int((y - 40) / 16) - 1                                'calculate which board the player is clicking
        col = Int((x + 4) / 16) - 1
        
        If bLButtonDown Then
           DrawNineBoard col, ln                             'if left click before right click, start auto mine sweep
        End If
    
        If arrBomb(col, ln).board = COMMON Then             'trun board into flagged board and bomb number - 1
          arrBomb(col, ln).board = FLAG
          DrawBoard col, ln, FLAG
          iLeftBombs = iLeftBombs - 1
          DrawNum iLeftBombs, 1
        ElseIf arrBomb(col, ln).board = FLAG Then           'trun flagged board into question mark board, and bomb number + 1
           arrBomb(col, ln).board = MARK
           DrawBoard col, ln, MARK
           iLeftBombs = iLeftBombs + 1
           DrawNum iLeftBombs, 1
        Else
            If arrBomb(col, ln).board = MARK Then
                arrBomb(col, ln).board = COMMON
                DrawBoard col, ln, COMMON
                                          
            End If
         End If
    End If
End Sub
  
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)    'when the player is moving mouse
    Dim ln As Integer, col As Integer
    Dim i As Integer, j As Integer
    
    If ((Not bLButtonDown) And (Not bRButtonDown)) Then                     'if the mouse is not clicked, exit sub
    Exit Sub
    End If
    
    
     If bDrawFace Then                       'if the player is holding left click on the face, change the picture of the face
     If x >= ptFace.x And y >= ptFace.y And x <= ptFace.x + 24 And y <= ptFace.y + 24 Then
        DrawFace SMILEDOWN                                          '3d special effect, change the face picture
      Else
      DrawFace iFace           'if the mouse if not on the face, exit sub
      Exit Sub
      End If
    End If
 
    
    If x < 12 Or x >= 12 + 16 * iCols Or y < 55 Or y >= 55 + 16 * iRows Then   'if the mouse is not above the board
      For i = iOldCol - 2 To iOldCol + 2
        For j = iOldRow - 2 To iOldRow + 2
          If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then         'prevernt mouse from moving out of game boundries
            DrawBoard i, j, arrBomb(i, j).board
          End If
        Next
      Next
      iOldRow = 0
      iOldCol = 0
      Exit Sub
    End If
                              '
    
    ln = Int((y - 40) / 16) - 1
    col = Int((x + 4) / 16) - 1

    
    If (ln = iOldRow And col = iOldCol) Then Exit Sub
    DrawBoard iOldCol, iOldRow, arrBomb(iOldCol, iOldRow).board
    If bLButtonDown And bRButtonDown Then
        For i = iOldCol - 2 To iOldCol + 2
            For j = iOldRow - 2 To iOldRow + 2
              If Not ((i < 0) Or (j < 0) Or (i >= iCols) Or (j >= iRows)) Then
                DrawBoard i, j, arrBomb(i, j).board
              End If                             '
            Next
        Next
        If (Not (col < 0) Or (ln < 0) Or (col >= iCols) Or (ln >= iRows)) Then
         DrawNineBoard col, ln
        End If
    ElseIf bLButtonDown Then
        If ln < 0 Then Exit Sub
        If arrBomb(col, ln).board = COMMON Then DrawBoard col, ln, DOWN
        If arrBomb(col, ln).board = MARK Then DrawBoard col, ln, MARKDOWN
    End If
    iOldRow = ln
    iOldCol = col
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)   'when the player stops clicking the mouse
   Dim ln As Integer, col As Integer
    Dim sumBomb  As Integer, sumFlag As Integer
    
Dim j As Integer, lie As Integer, hang As Integer


     ln = Int((y - 40) / 16) - 1
        col = Int((x + 4) / 16) - 1
        
                'the new feature I add, make sure that there is no bomb in first click on board
                ' Here, I initilize the game after clicking
        
        Do While shumu < iBombs
           lie = Int(Rnd * iCols)
          hang = Int(Rnd * iRows)
          If arrBomb(lie, hang).isBomb = False And (Abs(hang - ln) > 1 Or Abs(lie - col) > 1) Then
              arrBomb(lie, hang).isBomb = True
             shumu = shumu + 1
           End If
      Loop
    

    
    If Button = 1 Then
        DrawFace iFace
       
        If x >= ptFace.x And y >= ptFace.y And x <= ptFace.x + 24 And y <= ptFace.y + 24 And bDrawFace Then
            OnGameNew
            Form_Paint
            Exit Sub
        End If                                              'if left click on face and then stop clicking, restart the game
        
        If bDrawFace Then
         bDrawFace = False
         bLButtonDown = False
         Exit Sub
        End If
        
        bDrawFace = False
        
        If ((x < 12) Or (x > 12 + 16 * iCols) Or (y < 55) Or (y > 55 + 16 * iRows)) Then
          bLButtonDown = False
          Exit Sub
        End If
        
        If iFace = CRY Or iFace = COOL Then Exit Sub                               'if the game is already over, exit sub
        
        ln = Int((y - 40) / 16) - 1
        col = Int((x + 4) / 16) - 1
        
        If ((col < 0) Or (ln < 0) Or (col >= iCols) Or (ln >= iRows)) Then Exit Sub           'prevent out-of-boundary click
        If Not bStarted Then
           bStarted = True
           Timer.Enabled = True                              'start timer
        End If
        
        If bRButtonDown Then
            sumBomb = 15 - arrBomb(col, ln).board
            sumFlag = 0
            For i = col - 1 To col + 1
               For j = ln - 1 To ln + 1
                If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then
                   If arrBomb(i, j).board = FLAG Then
                    sumFlag = sumFlag + 1                 'count the number of flags
                   End If
                End If
                Next
            Next
            If sumBomb = sumFlag Then
                For i = col - 1 To col + 1
                    For j = ln - 1 To ln + 1
                       If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then
                       If arrBomb(i, j).board = COMMON Or arrBomb(i, j).board = MARK Then
                        Kick i, j                              'if bomb number = flag number, check if the positions are right
                       End If
                       End If
                    Next
                Next
            End If
            
            For i = col - 1 To col + 1
               For j = ln - 1 To ln + 1
                   If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then              'prevent out-of-boundry click
                    DrawBoard i, j, arrBomb(i, j).board
                   End If
               Next
            Next
       
        ElseIf (((arrBomb(col, ln).board = COMMON) Or (arrBomb(col, ln).board = MARK)) And bLButtonDown) Then
               Kick col, ln
        End If
        bLButtonDown = False
        If iFace <> CRY Then Check
        
    ElseIf Button = 2 Then
        If iFace = CRY Or iFace = COOL Then Exit Sub
        
        bRButtonDown = False
        bLButtonDown = False
        If x < 12 Or x >= 12 + 16 * iCols Or y < 55 Or y >= 55 + 16 * iRows Then Exit Sub
        
        ln = Int((y - 40) / 16) - 1
        col = Int((x + 4) / 16) - 1
        
        If ((col < 0) Or (ln < 0) Or (col >= iCols) Or (ln >= iRows)) Then Exit Sub
     
     
     
     
     
     'the different action of stop left click or right clicking first after double clicking
        If bLButtonDown Then
            sumBomb = 15 - arrBomb(col, ln).board
            sumFlag = 0
            For i = col - 1 To col + 1
                For j = ln - 1 To ln + 1
                If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then
                If arrBomb(i, j).board = FLAG Then
                    sumFlag = sumFlag + 1                                   'calculate flag numbers
                End If
                End If
                Next
            Next
            If sumBomb = sumFlag Then
                For i = col - 1 To col + 1
                    For j = ln - 1 To ln + 1
                        If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then
                        If arrBomb(i, j).board = COMMON Or arrBomb(i, j).board = MARK Then
                          Kick i, j
                        End If
                        End If
                    Next
                Next
            End If
            For i = col - 1 To col + 1
                For j = ln - 1 To ln + 1
                   If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then
                    DrawBoard i, j, arrBomb(i, j).board
                   End If
                Next
            Next
        Else
        End If
        
        If iFace <> CRY Then Check
    End If
End Sub

Public Sub Form_Paint()
    Dim i As Integer, j As Integer
    Dim w As Long, h As Long
    Dim oldForecolor, oldFillcolor As ColorConstants
    
     w = ScaleWidth - 1     'scaleleft scalewidth scalewidth scaleheight is "drawboard" property. The defult settings - left top=0   width，height, and form are the same with that of the loaded picture
     h = ScaleHeight - 1
    
    oldForecolor = ForeColor
    oldFillcolor = FillColor
    
    FillColor = RGB(198, 195, 198)      'Set internal fill color and border color
    ForeColor = RGB(198, 195, 198)
    FillStyle = 0
    
    
    
    Line (0, 0)-(w, 550), , BF
    
    Line (0, 550)-(120, (h - 120)), , BF
    
    Line (0, (h - 120))-(w, h), , BF
    
    Line ((w - 110), 550)-(w, (h - 120)), , BF

    
    
    ForeColor = RGB(255, 255, 255)      'set fill color: white. Here, the below and right border color is drew
    DrawWidth = 1
    Line (w - 1, 0)-(0, 0)
    Line (0, 0)-(0, h)
    Line (w - 2, 1)-(1, 1)
    Line (1, 1)-(1, h - 1)
    Line (w - 3, 2)-(2, 2)
    Line (2, 2)-(2, h - 2)
    Line (w - 9, 10)-(w - 9, 45)
    Line (w - 9, 45)-(9, 45)
    Line (w - 10, 11)-(w - 10, 44)
    Line (w - 10, 44)-(10, 44)
    Line (w - 9, 53)-(w - 9, h - 9)
    Line (w - 9, h - 9)-(9, h - 9)
    Line (w - 10, 54)-(w - 10, h - 10)
    Line (w - 10, h - 10)-(10, h - 10)
    Line (w - 11, 55)-(w - 11, h - 11)
    Line (w - 11, h - 11)-(11, h - 11)
    
    
    ForeColor = RGB(132, 130, 132)   'set fill color: gray. Here, the upper and left border color is drew
    DrawWidth = 1
    Line (w, 1)-(w, h)
    Line (w, h)-(1, h)
    Line (w - 1, 1)-(w - 1, h - 1)
    Line (w - 1, h - 1)-(2, h - 1)
    Line (w - 2, 2)-(w - 2, h - 2)
    Line (w - 2, h - 2)-(3, h - 2)
    Line (w - 10, 9)-(9, 9)
    Line (9, 9)-(9, 45)
    Line (w - 11, 10)-(10, 10)
    Line (10, 10)-(10, 44)
    Line (w - 10, 52)-(9, 52)
    Line (9, 52)-(9, h - 9)
    Line (w - 11, 53)-(10, 53)
    Line (10, 53)-(10, h - 10)
    Line (w - 12, 54)-(11, 54)
    Line (11, 54)-(11, h - 11)
    
    
    FillColor = RGB(255, 255, 255)
    ForeColor = RGB(132, 130, 132)
    
    Line (ptFace.x - 1, ptFace.y - 1)-(ptFace.x + 25, ptFace.y + 25) 'set a line in the face icon
    
    
    ForeColor = oldForecolor
    FillColor = oldFillcolor
    
    DrawNum iLeftBombs, 1       'draw num of bombs left to be find
    
    DrawNum iTime, 2            'draw timer

    
    DrawFace iFace              'draw face
    
    
    For i = 0 To iCols - 1
      For j = 0 To iRows - 1
        DrawBoard i, j, arrBomb(i, j).board                 'draw board
      Next
    Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveRecord
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1, Me
End Sub

Private Sub mnuBegin_Click()

    mnuBegin.Checked = True
    mnuMiddle.Checked = False
    mnuExpert.Checked = False
    mnuCust.Checked = False
    
    iCols = 8
    iRows = 8
    iBombs = 10
    iLevel = 0          'default difficult level is "easy"
    OnGameNew
    Form_Paint
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuMiddle_Click()
    mnuBegin.Checked = False
    mnuMiddle.Checked = True
     mnuCust.Checked = False
    mnuExpert.Checked = False
    
    iCols = 16
    iRows = 16
    iBombs = 40
    iLevel = 1
    OnGameNew
    Form_Paint
End Sub

Private Sub mnuExpert_Click()
    mnuBegin.Checked = False
    mnuMiddle.Checked = False
    mnuExpert.Checked = True
    mnuCust.Checked = False
    
    iCols = 30
    iRows = 16
    iBombs = 99
    iLevel = 2
    OnGameNew
    Form_Paint
End Sub
Private Sub mnuCust_Click()
    frmCustom.Show 1, Me
End Sub


Private Sub mnuStart_Click()
    OnGameNew
    Form_Paint
End Sub

Private Sub mnuRecord_Click()   'show leaderboard form
    frmRecord.Show 1, Me
End Sub

Private Sub Timer_Timer()   'show playing time
    If iTime < 999 Then
        iTime = iTime + 1
        DrawNum iTime, 2
    End If
End Sub
Private Sub ReadRecord()   'read leaderboard record
    Dim i As Integer
    Dim v As Variant
    Dim s As String
        
    v = GetAllSettings("Mine", "Records")
    If IsEmpty(v) Then
        For i = 0 To 2
            username(i) = "unknown"
            userscore(i) = 999
        Next
        Exit Sub
    End If
    For i = 0 To 2
        If i >= LBound(v, 1) And i <= UBound(v, 1) Then
            username(i) = v(i, 0)
            userscore(i) = v(i, 1)
        Else
            username(i) = "unknown"
            userscore(i) = 999
        End If
    Next
End Sub
Private Sub SaveRecord()         'save leaderboard record
    Dim i As Integer
    Dim v As Variant
    
    v = GetAllSettings("Mine", "Records")
    If Not IsEmpty(v) Then
        DeleteSetting "Mine", "Records"   'delete records
    End If
    For i = 0 To 2
        SaveSetting "Mine", "Records", username(i), userscore(i)
    Next
End Sub

