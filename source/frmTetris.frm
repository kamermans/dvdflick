VERSION 5.00
Begin VB.Form frmTetris 
   Caption         =   "Tetris"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4650
   ControlBox      =   0   'False
   Icon            =   "frmTetris.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tGame 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   3840
      Top             =   3720
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   5700
      Width           =   1155
   End
   Begin VB.CommandButton cmdScores 
      Caption         =   "Highscores"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   2115
   End
   Begin VB.CommandButton cmdChicken 
      Caption         =   "I give up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   6240
      Width           =   2115
   End
   Begin VB.PictureBox picNext 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   3360
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   960
   End
   Begin VB.PictureBox picField 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   120
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   208
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   3120
   End
   Begin VB.Label Lines 
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   10
      Top             =   2580
      Width           =   1155
   End
   Begin VB.Label lblLines 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   9
      Top             =   2820
      Width           =   1155
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   7
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Label lblLevel 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   1500
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   1260
      Width           =   1155
   End
End
Attribute VB_Name = "frmTetris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Binary
Option Base 0


Private Type highScoreData
    Name As String
    Score As Long
End Type


' Block
Private Const BLOCK_SIZE As Long = 3
Private Const BLOCK_COUNT As Long = 6

' Field dimensions
Private Const fieldWidth As Long = 12
Private Const fieldHeight As Long = 24

' Field
Private Field(fieldWidth, fieldHeight) As Long
Private fieldColor As Long

' Block arrays
Private Blocks(BLOCK_COUNT, BLOCK_SIZE, BLOCK_SIZE) As Byte
Private Block(BLOCK_SIZE) As String

' Block info
Private blockX As Long
Private blockY As Long

Private curBlock(BLOCK_SIZE, BLOCK_SIZE) As Byte
Private curBlockCol As Long
Private blockWidth As Long
Private blockHeight As Long

Private nextBlock As Long
Private nextBlockCol As Long

' Misc
Private enableInput As Boolean
Private gameLoop As Boolean
Private endTetris As Boolean
Private gameDelay As Long
Private Score As Long
Private Level As Long
Private levelLines As Long

' Highscores
Private highScores(9) As highScoreData


Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Sub cmdChicken_Click()

    If cmdChicken.Caption = "Again!" Then
        cmdChicken.Caption = "I give up"
        startGame
        Exit Sub
    End If
    
    gameLoop = False

End Sub


Private Sub cmdExit_Click()

    gameLoop = False
    endTetris = True
    Me.Hide

End Sub


Private Sub cmdScores_Click()

    Dim A As Long
    Dim Scores As String
    Dim thisName As String
    
    
    Scores = "Highscores:" & vbNewLine & vbNewLine
    
    For A = 0 To 9
        thisName = highScores(A).Name & Space$(32 - Len(highScores(A).Name))
        Scores = Scores & thisName & vbTab & highScores(A).Score & vbNewLine
    Next A
    
    MsgBox Scores

End Sub


Private Sub Form_Activate()

    startGame

End Sub


Private Sub startGame()

    Dim Idx As Long
    
    
    Fill 0, 0, fieldWidth, fieldHeight, fieldColor
    
    Randomize Timer

    nextBlock = Int(Rnd * (BLOCK_COUNT + 1))
    Idx = Int(Rnd * 64) + 96
    nextBlockCol = RGB(Idx * 1.6, Idx * 1.2, 0)
    newBlock
    drawBlock
    
    gameDelay = 350
    Score = 0
    Level = 0
    levelLines = 0
    updateScoreboard
    
    
    gameLoop = True
    enableInput = True
    renderField
    
    tGame.Interval = gameDelay
    tGame.Enabled = True

End Sub


Private Sub gameOver()

    Dim A As Long, B As Long
    
    
    For A = 0 To 9
        If Score > highScores(A).Score Then
            
            For B = 9 To A + 1 Step -1
                highScores(B) = highScores(B - 1)
            Next B
            
            highScores(A).Name = InputBox("You've made it into the highscore list! Please enter your name.", "Highscore")
            highScores(A).Name = highScores(A).Name & Space$(32 - Len(highScores(A).Name))
            highScores(A).Score = Score
            Exit For
            
        End If
    Next A

    If Not endTetris Then
        cmdScores_Click
        cmdChicken.Caption = "Again!"
    Else
        endTetris = False
    End If
    
End Sub


Private Sub renderField()

    Dim X As Long, Y As Long
    

    picField.Cls
    For X = 0 To fieldWidth
        For Y = 0 To fieldHeight
            picField.Line (X * 16, Y * 16)-(X * 16 + 15, Y * 16 + 15), Field(X, Y), BF
        Next Y
        picField.Line (X * 16, 0)-(X * 16, fieldHeight * 16 + 15), RGB(63, 63, 63)
    Next X
    
    For Y = 0 To fieldHeight
        picField.Line (0, Y * 16)-(fieldWidth * 16 + 15, Y * 16), RGB(63, 63, 63)
    Next Y

End Sub


Private Sub dropBlock()

    Do: Loop Until moveBlock(0, 1) = False
    renderField

End Sub


Private Sub updateScoreboard()

    lblLevel.Caption = Level
    lblScore.Caption = Score
    lblLines.Caption = levelLines
    
End Sub


Private Sub Fill(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Col As Long)

    Dim X As Long
    Dim Y As Long
    
    
    For X = X1 To X2
        For Y = Y1 To Y2
            Field(X, Y) = Col
        Next Y
    Next X

End Sub


Private Sub drawNextBlock()

    Dim X As Long, Y As Long
    Dim Col As Long
    

    picNext.Cls
    For X = 0 To BLOCK_SIZE
        For Y = 0 To BLOCK_SIZE
            If Blocks(nextBlock, X, Y) = 0 Then
                Col = fieldColor
            Else
                Col = nextBlockCol
            End If
            
            picNext.Line (X * 16, Y * 16)-(X * 16 + 15, Y * 16 + 15), Col, BF
        Next Y
    Next X

End Sub


Private Sub Delay(ByVal MS As Long)

    Dim timeStart As Long
    

    timeStart = GetTickCount
    While GetTickCount - timeStart < MS
        DoEvents
    Wend

End Sub


Private Sub newBlock()

    Dim X As Long
    Dim Y As Long
    Dim Idx As Long
    

    curBlockCol = nextBlockCol
    
    blockWidth = 0
    blockHeight = 0
    
    For X = 0 To BLOCK_SIZE
        For Y = 0 To BLOCK_SIZE
            
            If Blocks(nextBlock, X, Y) <> 0 Then
                If X > blockWidth Then blockWidth = X
                If Y > blockHeight Then blockHeight = Y
            End If
            
            curBlock(X, Y) = Blocks(nextBlock, X, Y)
            
        Next Y
    Next X

    blockX = (fieldWidth \ 2) - (blockWidth \ 2)
    blockY = 0
    
    nextBlock = Int(Rnd * (BLOCK_COUNT + 1))
    Idx = Int(Rnd * 64) + 96
    nextBlockCol = RGB(Idx * 1.6, Idx * 1.2, 0)
    drawNextBlock

End Sub


Private Sub drawBlock(Optional ByVal Remove As Boolean = False)

    Dim X As Long, Y As Long
    Dim Col As Long
    
    
    If Remove Then
        Col = fieldColor
    Else
        Col = curBlockCol
    End If
    
    For X = 0 To blockWidth
        For Y = 0 To blockHeight
            If curBlock(X, Y) = 1 Then Field(X + blockX, Y + blockY) = Col
        Next Y
    Next X

End Sub


Private Sub rotateBlock()

    Dim X As Long, Y As Long
    Dim xStart As Long
    Dim yStart As Long
    Dim tBlock(BLOCK_SIZE, BLOCK_SIZE) As Byte
    
    
    drawBlock True
    
    Do
    
        For X = 0 To BLOCK_SIZE
            For Y = 0 To BLOCK_SIZE
                tBlock(BLOCK_SIZE - Y, X) = curBlock(X, Y)
            Next Y
        Next X
        
        
        ' Find top left origin
        xStart = BLOCK_SIZE
        yStart = BLOCK_SIZE
        
        For X = 0 To BLOCK_SIZE
            For Y = 0 To BLOCK_SIZE
                If tBlock(X, Y) <> 0 Then
                    If X < xStart Then xStart = X
                    If Y < yStart Then yStart = Y
                End If
                
                ' Nullify old block
                curBlock(X, Y) = 0
            Next Y
        Next X
        
        
        ' Copy back
        blockWidth = 0
        blockHeight = 0
        
        For X = xStart To BLOCK_SIZE
            For Y = yStart To BLOCK_SIZE
                curBlock(X - xStart, Y - yStart) = tBlock(X, Y)
                
                If curBlock(X - xStart, Y - yStart) <> 0 Then
                    If X - xStart > blockWidth Then blockWidth = X - xStart
                    If Y - yStart > blockHeight Then blockHeight = Y - yStart
                End If
            Next Y
        Next X
    
    Loop Until blockCollision = False
    
    drawBlock False

End Sub


Private Function moveBlock(ByVal X As Long, ByVal Y As Long) As Boolean

    moveBlock = True
    enableInput = False
    drawBlock True
    
    blockX = blockX + X
    If blockCollision Then
        blockX = blockX - X
        moveBlock = False
    End If
    
    blockY = blockY + Y
    If blockCollision Then
        blockY = blockY - Y
        drawBlock
        
        scanLines blockY, blockY + blockHeight
        
        newBlock
        If blockCollision Then
            gameLoop = False
        Else
            drawBlock
        End If
        
        moveBlock = False
        enableInput = True
        Exit Function
    End If
        
    drawBlock
    enableInput = True

End Function


Private Sub scanLines(ByVal startRow As Long, ByVal endRow As Long)

    Dim A As Long
    Dim Y As Long, X As Long
    Dim Full As Boolean
    Dim Consec As Long
    
    
    For Y = startRow To endRow
        
        Full = True
        For X = 0 To fieldWidth
            If Field(X, Y) = fieldColor Then Full = False
        Next X
        
        If Full = True Then
            For A = 0 To 3
                Fill 0, Y, fieldWidth, Y, vbWhite
                renderField
                Delay 40
                Fill 0, Y, fieldWidth, Y, fieldColor
                renderField
                Delay 40
            Next A
            
            
            levelLines = levelLines + 1
            If levelLines >= 11 Then
                Level = Level + 1
                gameDelay = gameDelay - 25
                If gameDelay < 25 Then gameDelay = 25
                levelLines = 0
            End If
            
            Consec = Consec + 1
            If Consec = 1 Then Score = Score + ((Level + 1) * 40)
            If Consec = 2 Then Score = Score + ((Level + 1) * 100)
            If Consec = 3 Then Score = Score + ((Level + 1) * 300)
            If Consec = 4 Then Score = Score + ((Level + 1) * 1200)
            
            updateScoreboard
            lowerField Y
            renderField
        End If
        
    Next Y

End Sub


Private Sub lowerField(ByVal Row As Long)

    Dim X As Long
    Dim Y As Long
    
    
    For Y = Row To 1 Step -1
        For X = 0 To fieldWidth
            Field(X, Y) = Field(X, Y - 1)
        Next X
    Next Y

End Sub


Private Function blockCollision() As Boolean

    Dim X As Long
    Dim Y As Long
    
    
    blockCollision = False
    
    For X = 0 To blockWidth
        For Y = 0 To blockHeight
        
            If curBlock(X, Y) = 1 Then
                If X + blockX < 0 Then
                    blockCollision = True
                    Exit Function
                
                ElseIf X + blockX > fieldWidth Then
                    blockCollision = True
                    Exit Function
                End If
                
                
                If Y + blockY < 0 Then
                    blockCollision = True
                    Exit Function
                
                ElseIf Y + blockY > fieldHeight Then
                    blockCollision = True
                    Exit Function
                    
                End If
                
                
                If Field(X + blockX, Y + blockY) <> fieldColor Then
                    blockCollision = True
                    Exit Function
                End If
            End If
        
        Next Y
    Next X

End Function


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Const keyUp = 38
    Const keyDown = 40
    Const keyLeft = 37
    Const keyRight = 39
    Const keyDrop = 32

    
    If enableInput = False Then Exit Sub
   
    Select Case KeyCode
        Case keyLeft
            moveBlock -1, 0
            renderField
        Case keyRight
            moveBlock 1, 0
            renderField
        Case keyDown
            moveBlock 0, 1
            renderField
        Case keyUp
            rotateBlock
            renderField
        Case keyDrop
            dropBlock
            renderField
    End Select

End Sub


Private Sub Form_Load()

    Dim A As Long
    Dim fileObj As clsBinaryFile


    appLog.Add "Loading " & Me.Name & "..."
    

    Block(0) = "0100"
    Block(1) = "0100"
    Block(2) = "1100"
    Block(3) = "0000"
    setBlock 0
    
    Block(0) = "0100"
    Block(1) = "1110"
    Block(2) = "0000"
    Block(3) = "0000"
    setBlock 1
    
    Block(0) = "1000"
    Block(1) = "1000"
    Block(2) = "1000"
    Block(3) = "1000"
    setBlock 2
    
    Block(0) = "1100"
    Block(1) = "1100"
    Block(2) = "0000"
    Block(3) = "0000"
    setBlock 3
    
    Block(0) = "1100"
    Block(1) = "0110"
    Block(2) = "0000"
    Block(3) = "0000"
    setBlock 4
    
    Block(0) = "0110"
    Block(1) = "1100"
    Block(2) = "0000"
    Block(3) = "0000"
    setBlock 5
    
    Block(0) = "1000"
    Block(1) = "1000"
    Block(2) = "1100"
    Block(3) = "0000"
    setBlock 6
    
    
    Set fileObj = New clsBinaryFile
    fileObj.fileOpen DATA_PATH & "tetris.dat", False

    ' Reset scores
    ' These appear in no particular order!
    If fileObj.fileLength = 0 Then
        highScores(0).Name = "kbosward"
        highScores(0).Score = "10000"
        
        highScores(1).Name = "Ruler"
        highScores(1).Score = "9000"
        
        highScores(2).Name = "WaltP"
        highScores(2).Score = "8000"
        
        highScores(3).Name = "Guzeppi"
        highScores(3).Score = "7000"
        
        highScores(4).Name = "Neal"
        highScores(4).Score = "6000"
        
        highScores(5).Name = "FFmpeg"
        highScores(5).Score = "5000"
        
        highScores(6).Name = "DVDAuthor"
        highScores(6).Score = "4000"
        
        highScores(7).Name = "ImgBurn"
        highScores(7).Score = "3000"
        
        highScores(8).Name = "Butter"
        highScores(8).Score = "3"
        
        highScores(9).Name = "Exl"
        highScores(9).Score = "1"
        
    ' Read scores
    Else
        For A = 0 To 9
            highScores(A).Name = fileObj.readString
            highScores(A).Score = fileObj.readLong
        Next A

    End If

    Set fileObj = Nothing
    
End Sub


Private Sub setBlock(ByVal Index As Long)

    Dim A As Long, B As Long
    Dim Char As String * 1
    
    
    For A = 0 To BLOCK_SIZE
        For B = 1 To BLOCK_SIZE + 1
            Char = Mid$(Block(A), B, 1)
            If Char = "1" Then Blocks(Index, B - 1, A) = 1
        Next B
    Next A
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim A As Long
    Dim fileObj As clsBinaryFile
    
    
    Set fileObj = New clsBinaryFile
    fileObj.fileOpen DATA_PATH & "tetris.dat", True
    
    For A = 0 To 9
        fileObj.writeString highScores(A).Name
        fileObj.writeLong highScores(A).Score
    Next A

    Set fileObj = Nothing

End Sub


Private Sub tGame_Timer()

    tGame.Interval = gameDelay
    
    If enableInput = True Then
    
        enableInput = False
        
        moveBlock 0, 1
        renderField
        
        enableInput = True
        
    End If
    
    If gameLoop = False Then
        tGame.Enabled = False
        gameOver
    End If

End Sub
