Attribute VB_Name = "Module"
Option Explicit

'define the possible condition of (below) board
Enum Board_Status
    COMMON = 0
    FLAG = 1
    MARK = 2
    DIEBOMB = 3
    NOBOMB = 4
    BOMB = 5
    MARKDOWN = 6
    NUM8 = 7
    NUM7 = 8
    NUM6 = 9
    NUM5 = 10
    NUM4 = 11
    NUM3 = 12
    NUM2 = 13
    NUM1 = 14
    DOWN = 15
End Enum

'define the possible condition of face
Enum Face_Status
    SMILEDOWN = 0
    COOL = 1
    CRY = 2
    SURPRISE = 3
    SMILE = 4
End Enum

'define the possible condition of (upper) board
Type TBOMB
    isBomb As Boolean
    board As Board_Status
End Type

Type TPOINT                         'define point xy position
    x As Integer
    y As Integer
End Type

Public picBoard(0 To 15) As StdPicture
Public picNum(0 To 10) As StdPicture
Public picFace(0 To 4) As StdPicture

Public ptFace As TPOINT                     'save face position
Public arrBomb() As TBOMB               'save bomb condition
Public bDrawFace As Boolean                 'whether to draw face
Public iFace As Integer                     'face condition

Public iBombs As Integer                    'total num of bombs
Public iLeftBombs As Integer                'remaining num of bombs

Public iRows As Integer, iCols As Integer   'rows and cols of board area

Public iLevel As Integer                    'difficult level:1¡«4£¨correspond to easy, medium, hard and custom£©
Public iTime As Integer             'best time record

Public bMark As Boolean

Public iOldRow As Integer, iOldCol As Integer                   'the row and col of the precious game

Public bStarted As Boolean                  'start icon

Public bLButtonDown As Boolean, bRButtonDown As Boolean
Public username(0 To 2) As String, userscore(0 To 2) As Integer         'leaderboard record

Public shumu As Integer


Sub DrawBoard(col As Integer, ln As Integer, board As Board_Status)
    frmMain.PaintPicture picBoard(board), (col + 1) * 16 - 4, 40 + (ln + 1) * 16, 16, 16

     
End Sub


Sub DrawNum(inum As Integer, itype As Integer)

    Dim x As Integer, i As Integer
    Dim strNum, strChr As String
    
    If itype = 1 Then x = 17
    If itype = 2 Then x = frmMain.Width / Screen.TwipsPerPixelX - 60
    
    
    
    
    If inum < 0 Then
        strNum = Format(inum, "00")
    Else
        strNum = Format(inum, "000")
    End If
    
    For i = 1 To 3
        strChr = Mid(strNum, i, 1)
        If strChr >= "0" And strChr <= "9" Then
            frmMain.PaintPicture picNum(Val(strChr)), x, 15, 13, 26
        Else
            frmMain.PaintPicture picNum(10), x, 15, 13, 26
        End If
        x = x + 13
    Next

End Sub

Sub DrawNineBoard(col As Integer, ln As Integer)
    Dim i As Integer, j As Integer
    For i = col - 1 To col + 1
        For j = ln - 1 To ln + 1
            If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then
                If arrBomb(i, j).board = COMMON Then DrawBoard i, j, DOWN
                If arrBomb(i, j).board = MARK Then DrawBoard i, j, MARKDOWN
            End If
        Next
    Next
End Sub

Sub DrawFace(face As Integer)
    frmMain.PaintPicture picFace(face), ptFace.x, ptFace.y, 24, 24
End Sub


Sub GameOver()
    Dim i As Integer, j As Integer
    For i = 0 To iCols - 1
        For j = 0 To iRows - 1
            If arrBomb(i, j).isBomb And (Not arrBomb(i, j).board = FLAG) Then
               arrBomb(i, j).board = BOMB
                DrawBoard i, j, BOMB
            End If
            If (Not arrBomb(i, j).isBomb) And arrBomb(i, j).board = FLAG Then
                arrBomb(i, j).board = NOBOMB
                DrawBoard i, j, NOBOMB
            End If
        Next
    Next
    
    iFace = CRY
    DrawFace iFace
    
    frmMain.Timer.Enabled = False
End Sub


Sub OnGameNew()


    Dim i As Integer, j As Integer
    Dim ln As Integer, col As Integer
    
    frmMain.Timer.Enabled = False
    
    Select Case (iLevel)
      Case 0
        iCols = 9
        iRows = 9
        iBombs = 10
      Case 1
        iCols = 16
        iRows = 16
        iBombs = 40
      Case 2
        iCols = 30
        iRows = 16
        iBombs = 99
    End Select
    
    ReDim arrBomb(0 To iCols, 0 To iRows)
    
    frmMain.Width = 30 * Screen.TwipsPerPixelX + 16 * Screen.TwipsPerPixelX * iCols

    frmMain.Height = 100 * Screen.TwipsPerPixelY + 13 * Screen.TwipsPerPixelY + 16 * Screen.TwipsPerPixelY * iRows + 4 * Screen.TwipsPerPixelY
    
                
    bLButtonDown = False
    bRButtonDown = False
    bStarted = False
    bDrawFace = False
    iFace = SMILE
    iOldCol = 0
    iOldRow = 0
    iLeftBombs = iBombs
    
    For i = 0 To iCols - 1
        For j = 0 To iRows - 1
            arrBomb(i, j).isBomb = False
            arrBomb(i, j).board = COMMON
        Next
    Next
    
    i = 0
    
      
    
    ptFace.x = frmMain.Width / (2 * Screen.TwipsPerPixelX) - 13
    ptFace.y = 15
    
    iTime = 0
    shumu = 0

End Sub

Sub Check()
     Dim WIN As Boolean
     Dim i, j As Integer
     
     WIN = True
    
      For i = 0 To iCols - 1
       For j = 0 To iRows - 1
        If arrBomb(i, j).board = COMMON Or arrBomb(i, j).board = MARK Then WIN = False
        If arrBomb(i, j).isBomb = False And arrBomb(i, j).board = FLAG Then WIN = False
      Next
     Next
     
      If Not WIN Then
        Exit Sub
      End If
      
     iFace = COOL
     DrawFace iFace
     frmMain.Timer.Enabled = False
     
     If iLevel < 3 Then
        If userscore(iLevel) > iTime Then
            userscore(iLevel) = iTime
            username(iLevel) = InputBox("Congratulations, you break the record of this level!" & Chr(10) & Chr(13) & "your name£º", "MineSweeper", "Unknown")
            If username(iLevel + 1) = "" Then username(iLevel + 1) = "Unknown"
        Else
            MsgBox "Congratulations, you win", vbOKOnly + vbInformation, "Minesweeper"
        End If
     Else
        MsgBox "Congratulations, you win", vbOKOnly + vbInformation, "Minesweeper"
     End If
 
End Sub


Sub Kick(col As Integer, ln As Integer)
 Dim sum As Integer
 Dim i As Integer, j As Integer
 
  
    If col < 0 Or ln < 0 Or col >= iCols Or ln >= iRows Then Exit Sub
    
    If arrBomb(col, ln).isBomb = True Then
      GameOver
      arrBomb(col, ln).board = DIEBOMB
      DrawBoard col, ln, DIEBOMB
      Exit Sub
    End If
    
    sum = 0
    i = 0
    j = 0
    For i = col - 1 To col + 1
      For j = ln - 1 To ln + 1
        If Not (i < 0 Or j < 0 Or i >= iCols Or j >= iRows) Then
            If arrBomb(i, j).isBomb Then
            sum = sum + 1
         End If
        End If
      Next
    Next
    
    arrBomb(col, ln).board = 15 - sum
    DrawBoard col, ln, 15 - sum
    
    If sum = 0 Then
      For i = col - 1 To col + 1
        For j = ln - 1 To ln + 1
        If i >= 0 And j >= 0 And i < iCols And j < iRows Then
          If ((Not (i = col And j = ln)) And (arrBomb(i, j).board = COMMON)) Then
             Kick i, j
          End If
        End If
        Next
      Next
    End If

End Sub
