VERSION 5.00
Begin VB.Form frmgame 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bricks"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10830
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrmove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   4080
   End
   Begin VB.Timer tmrbrickbreak 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10320
      Top             =   3240
   End
   Begin VB.Label lbllevel 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4680
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Shape Paddle 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   600
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label lbllives 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lbllivestitle 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Lives: "
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblquit 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Quit (esc)"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label lblscore 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblscoretitle 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Score: "
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label lblspace 
      BackColor       =   &H00000000&
      Caption         =   "Press Space To Begin"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   4440
      Width           =   5415
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   47
      Left            =   9480
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   46
      Left            =   8160
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   45
      Left            =   6840
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   44
      Left            =   5520
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   43
      Left            =   4200
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   42
      Left            =   2880
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   41
      Left            =   1560
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   40
      Left            =   240
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   39
      Left            =   9480
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   38
      Left            =   8160
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   37
      Left            =   6840
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   36
      Left            =   5520
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   35
      Left            =   4200
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   34
      Left            =   2880
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   33
      Left            =   1560
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   32
      Left            =   240
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   31
      Left            =   9480
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   30
      Left            =   8160
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   29
      Left            =   6840
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   28
      Left            =   5520
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   27
      Left            =   4200
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   26
      Left            =   2880
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   25
      Left            =   1560
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   24
      Left            =   240
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   23
      Left            =   9480
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   22
      Left            =   8160
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   21
      Left            =   6840
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   20
      Left            =   5520
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   19
      Left            =   4200
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   18
      Left            =   2880
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   17
      Left            =   1560
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   16
      Left            =   240
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   15
      Left            =   9480
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   14
      Left            =   8160
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   13
      Left            =   6840
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   12
      Left            =   5520
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   11
      Left            =   4200
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   10
      Left            =   2880
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   9
      Left            =   1560
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H00FFFF80&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   8
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Ball 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   255
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   7
      Left            =   9480
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   6
      Left            =   8160
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   5
      Left            =   6840
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   4
      Left            =   5520
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   4200
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   2880
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   1560
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Brick 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   0
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Shape Danger 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0FF&
      Height          =   975
      Left            =   -240
      Top             =   6480
      Width           =   13215
   End
End
Attribute VB_Name = "frmgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim level As Integer
Dim LevelSelect As String
Dim Lives As Integer
Dim Boundary As Integer
Dim xvel, yvel As Integer
Dim Done As Boolean
Dim Brickdisplay As Integer
Public playername As String
Public Multiplier As Integer
'** Declare Variables

Private Sub lose_life()

If Lives = 1 Then       'if the player has lost their last life
    message = MsgBox("You Lost All Of Your Lives, Game Over.", , "Game Over")   'display a message saying they've lost
    FileName = App.Path & "\scores.txt"     'assign the filename
    Open FileName For Random As #1 Len = Len(CurrentPlayer)         'open the file
    NumberOfRecords = LOF(1) / Len(CurrentPlayer)                   'calculate the number of records in the file
    Pointer = NumberOfRecords + 1                                   'make the pointer one ahead of the length of the file
    Put #1, Pointer, CurrentPlayer                                  'put the pointer and the current players information into the file
    Close #1                                                        'close the file
    NumberOfRecords = NumberOfRecords + 1                           'add 1 to the number of records (to include the current player!)
    
    End
    
'    '**hide the game form and show the home form
'    frmhome.Show
'    frmhome.Enabled = True
'    frmgame.Hide
Else
    Lives = Lives - 1
    lbllives.Caption = Lives
    Paddle.Left = 600
    lblspace.Visible = True
    '*** If the player still has lives, then take one life off and resume the game.
End If


End Sub
Private Sub Complete_game()

    message = MsgBox("You Beat The Game, Congratulations!", , "Winner!!")           'display a message
    FileName = App.Path & "\scores.txt"
    Open FileName For Random As #1 Len = Len(CurrentPlayer)
        NumberOfRecords = LOF(1) / Len(CurrentPlayer)                               'send the players information to file
        Pointer = NumberOfRecords + 1
        Put #1, Pointer, CurrentPlayer
    Close #1
    NumberOfRecords = NumberOfRecords + 1
    
    End

End Sub
Private Sub Form_Load()
    
    CurrentPlayer.score = 0
    Lives = 10
    level = 1
    '** initialise the score, lives and start on level 1
    
    playername = InputBox("Please Enter Your Name Or A Nickname That Is Less Than 15 Characters In Length.", "Player Name")
    If Len(playername) > 15 Or playername = "" Then
        Do
            playername = InputBox("Whoops, Please Try Entering Your Name Or A Nickname That Is 15 Character Or Less.", "Player Name")
        Loop Until Len(playername) <= 15 And playername <> ""
    End If
    'make the current player enter their name until its under 15 characters long and isnt blank
    
    
    LevelSelect = InputBox("Please Select A Level. Easy (1), Medium (2), Hard (3) ((Enter 1,2 or 3))", "Level Select")
    If LevelSelect <> "1" And LevelSelect <> "2" And LevelSelect <> "3" Then
        Do
            LevelSelect = InputBox("Whoops, Please Enter 1, 2 or 3. (Easy (1), Medium (2), Hard (3))", "Level Select")
        Loop Until LevelSelect = "1" Or LevelSelect = "2" Or LevelSelect = "3"
    End If
    '** make the player select a level from one of the three options
    
    Select Case LevelSelect
        Case Is = "1"
            Paddle.Width = 2415
            Boundary = 8520
            Multiplier = 1
        Case Is = "2"
            Paddle.Width = 1935
            Boundary = 9000
            Multiplier = 2
        Case Is = "3"
            Paddle.Width = 1335
            Boundary = 9600
            Multiplier = 3
    End Select
    '** depending on the level selected, set the relevant multipleir, boundary and size of the paddle
    

    lblscore.Caption = CurrentPlayer.score
    lbllives.Caption = Lives
    
    Call startlevel         'start the game
    
    End Sub

Private Sub lblquit_Click()

frmhome.Show
frmhome.Enabled = True
frmgame.Hide

'** hide this form (the game) and show the menu (home page)

End Sub

Public Sub startlevel()
Select Case level
    Case Is = 1
    Brickdisplay = 15
    CurrentPlayer.score = 0
    CurrentPlayer.playername = playername
    Case Is = 2
    Brickdisplay = 23
    Case Is = 3
    Brickdisplay = 31
    Case Is = 4
    Brickdisplay = 39
    Case Is = 5
    Brickdisplay = 47
    Case Is = 6
        Call Complete_game
End Select
'** select the correct amount of bricks depending on the level

For i = 0 To 47
    Brick(i).Visible = False
Next
'** make all the bricks invisible

For i = 0 To Brickdisplay
    Brick(i).Visible = True
Next
'** make the correct number of bricks visible

Paddle.Left = 600
Ball.Left = 0
lblspace.Visible = True
lbllevel.Visible = True
lbllevel.Caption = "Stage " & level

End Sub

Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case vbKeySpace
        lblspace.Visible = False
        lbllevel.Visible = False
        lbllives.Caption = Lives
        If tmrmove.Enabled = False Then     'this starts the game when the spacebar is pressed
            yvel = 40
            xvel = 40
            tmrmove.Enabled = True
            tmrbrickbreak.Enabled = True
        End If
    Case Is = 39
        If Paddle.Left >= Boundary Then
            Paddle.Left = Paddle.Left
        Else
        Paddle.Left = Paddle.Left + 200           'if the left arrow key (39 (ascii code)) is pressed, the paddle moves to the left
        End If
    Case Is = 37
        If Paddle.Left <= 0 Then
            Paddle.Left = Paddle.Left
        Else
            Paddle.Left = Paddle.Left - 200           'if the right arrow key (37 (ascii code)) is pressed,, the paddle moves to the right
        End If
    Case Is = 27
        frmgame.Hide
        frmgame.Enabled = False
        frmhome.Enabled = True                      'if the escape key (27 (ascii code)) is pressed, take the player to the home screen
        frmhome.Show
    End Select
    
End Sub
Private Sub flash_danger()

Danger.BackColor = &HFF&
lbllives.BackColor = &HFF&
lblscore.BackColor = &HFF&
lbllivestitle.BackColor = &HFF&
lblscoretitle.BackColor = &HFF&     'turn everything in the dangerzone red

Pause 0.2                               'wait 0.2 seconds

Danger.BackColor = &HC0C0FF
lbllives.BackColor = &HC0C0FF
lblscore.BackColor = &HC0C0FF
lbllivestitle.BackColor = &HC0C0FF      'return things in the danger zone are to original colour
lblscoretitle.BackColor = &HC0C0FF

Pause 0.2                               'wait 0.2 seconds

Danger.BackColor = &HFF&
lbllives.BackColor = &HFF&
lblscore.BackColor = &HFF&
lbllivestitle.BackColor = &HFF&         'turn everything in the dangerzone red
lblscoretitle.BackColor = &HFF&

Pause 0.2                               'wait 0.2 seconds

Danger.BackColor = &HC0C0FF
lbllives.BackColor = &HC0C0FF
lblscore.BackColor = &HC0C0FF           'turn everything in the dangerzone back to its original colour
lbllivestitle.BackColor = &HC0C0FF
lblscoretitle.BackColor = &HC0C0FF

End Sub
Sub Pause(ByVal nSecond As Single)
Dim t0 As Single
Dim dummy As Integer

    t0 = Timer
    Do While Timer - t0 < nSecond
    dummy = DoEvents()
    
    If Timer < t0 Then
        t0 = t0 - 24 * 60 * 60
        
    End If
    Loop
End Sub

Private Sub tmrmove_Timer()

If Ball.Left < 0 Then xvel = xvel * -1
If Ball.Left > Me.Width - Ball.Width Then xvel = xvel * -1
If Ball.Top < 0 Then yvel = yvel * -1                           'if the ball touches any surface, make the angle the opposite
If Ball.Top > Danger.Top Then
    Ball.Left = 0
    Ball.Top = 4560
    
    Call lose_life
    
    tmrmove.Enabled = False
    
    Exit Sub
End If

If Ball.Top = 6480 Then
    Call flash_danger
End If


If Ball.Left + Ball.Width > Paddle.Left And Ball.Left < Paddle.Left + Paddle.Width Then
    If Ball.Top + Ball.Height > Paddle.Top And Ball.Top < Paddle.Top + Paddle.Height Then
        If yvel > 0 Then
            yvel = yvel * -1
            xvel = xvel + (pan / 2)
        End If
    End If
End If

Ball.Left = Ball.Left + xvel

Ball.Top = Ball.Top + yvel

End Sub

Private Sub tmrbrickbreak_Timer()

'This controls the angles which the ball takes when it hits a block
For i = 0 To Brickdisplay
    If Brick(i).Visible = True Then
        If Ball.Left + Ball.Width > Brick(i).Left And Ball.Left < Brick(i).Left + Brick(i).Width Then
            If Ball.Top + Ball.Height > Brick(i).Top And Ball.Top < Brick(i).Top + Brick(i).Height Then
                Brick(i).Visible = False    'makes the block disappear when hit
                CurrentPlayer.score = CurrentPlayer.score + (10 * Multiplier)       'add to score
                lblscore.Caption = CurrentPlayer.score                              'displays the score
                yvel = yvel * -1                'makes the angle that the ball is moving opposite when a block is hit
            End If
        End If
    End If
Next i

Done = True                             'make done = true

For i = 0 To Brickdisplay
    If Brick(i).Visible = True Then Done = False            'if any bricks are still visible, then make done = false
Next

If Done = True Then
    Ball.Left = 0
    Ball.Top = 4680
    tmrmove.Enabled = False             'if done is still true, take the player to the next level
    level = level + 1

    Call startlevel
End If
End Sub
