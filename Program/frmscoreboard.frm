VERSION 5.00
Begin VB.Form frmscoreboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsearch 
      Caption         =   "Search"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtsearch 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Text            =   "Type Your Name Here To Search For Your Place!"
      Top             =   1320
      Width           =   3855
   End
   Begin VB.ListBox lstscores 
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5925
      Left            =   600
      TabIndex        =   0
      Top             =   1680
      Width           =   7095
   End
   Begin VB.Label lblquit 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit To Menu"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1335
      Left            =   6960
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblgame 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lbltitle 
      BackStyle       =   0  'Transparent
      Caption         =   "High Scores"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image imgbackground 
      Enabled         =   0   'False
      Height          =   8535
      Left            =   0
      Picture         =   "frmscoreboard.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmscoreboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Scores(199) As Integer
Dim Playername(199) As String

Private Sub cmdsearch_Click()

Call Search_name(Array_of_players())

End Sub

Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then Call Search_name(Array_of_players())

End Sub

Private Sub Search_name(ByRef Array_of_players() As PlayerType)
Dim Searchname As String

Searchname = txtsearch.Text

lstscores.Clear

For i = 0 To NumberOfRecords                    'for every record in the file
    If InStr(1, LCase(Array_of_players(i).Playername), LCase(Searchname), vbTextCompare) <> 0 Then          'compare the search name with the current name in the file
        lstscores.AddItem Array_of_players(i).Playername & vbTab & Array_of_players(i).score
    End If
Next i

If lstscores = "" Then
    lstscores.AddItem "Sorry, There are no more of this name in "
    lstscores.AddItem " our records."
    lstscores.AddItem "Try searching for another name!"
End If

'** search for the name that the user wishes in the scores file

End Sub

Private Sub Form_Load()

Call Load_scores

End Sub

Private Sub Load_scores()
Dim LengthOfName(200) As Integer

    FileName = App.Path & "\scores.txt"
    Open FileName For Random As #1 Len = Len(CurrentPlayer)
    NumberOfRecords = LOF(1) / Len(CurrentPlayer)
If CurrentPlayer.score > 0 Then
    Pointer = NumberOfRecords + 1
    Put #1, Pointer, CurrentPlayer
End If
    Close #1

ReDim Array_of_players(NumberOfRecords)

'Read in the file of scores
FileName = App.Path & "\scores.txt"

Open FileName For Random As #1 Len = Len(CurrentPlayer)
    For i = 1 To NumberOfRecords
        Get #1, i, CurrentPlayer
        Array_of_players(i - 1).Playername = CurrentPlayer.Playername
        Array_of_players(i - 1).score = CurrentPlayer.score
    Next i
Close #1

CurrentPlayer.Playername = frmgame.Playername
CurrentPlayer.score = 0

'** sort the scores into highest to lowest
For Outer = 0 To NumberOfRecords
    For Inner = Outer + 1 To NumberOfRecords
        If Array_of_players(Outer).score < Array_of_players(Inner).score Then
            TempScore = Array_of_players(Outer).score
            TempName = Array_of_players(Outer).Playername
            Array_of_players(Outer).score = Array_of_players(Inner).score
            Array_of_players(Outer).Playername = Array_of_players(Inner).Playername
            Array_of_players(Inner).score = TempScore
            Array_of_players(Inner).Playername = TempName
        End If
    Next Inner
Next Outer

lstscores.AddItem "Names: " & vbTab & vbTab & "Scores: "
lstscores.AddItem ""

For i = 0 To NumberOfRecords
    For j = 1 To 15
        If Mid(Array_of_players(i).Playername, j, 1) <> " " Then
            LengthOfName(i) = LengthOfName(i) + 1
        End If
    Next j
Next i

'** display all of the scores and the players name
For i = 0 To NumberOfRecords
    If LengthOfName(i) > 5 Then
        lstscores.AddItem Array_of_players(i).Playername & vbTab & Array_of_players(i).score
    Else
        lstscores.AddItem Array_of_players(i).Playername & vbTab & vbTab & Array_of_players(i).score
    End If
Next i

End Sub


Private Sub lblgame_Click()
frmscoreboard.Hide
frmscoreboard.Enabled = False
frmgame.Enabled = True
frmgame.Show
End Sub

Private Sub lblquit_Click()

frmscoreboard.Hide
frmscoreboard.Enabled = False
frmhome.Enabled = True
frmhome.Show

End Sub

Private Sub lbltitle_Click()

txtsearch.Text = ""
lstscores.Clear


lstscores.AddItem "Names: " & vbTab & vbTab & "Scores: "
lstscores.AddItem ""

Dim LengthOfName(200) As Integer

For i = 0 To NumberOfRecords
    For j = 1 To 15
        If Mid(Array_of_players(i).Playername, j, 1) <> " " Then
            LengthOfName(i) = LengthOfName(i) + 1
        End If
    Next j
Next i

'** display all of the scores and the players name
For i = 0 To NumberOfRecords
    If LengthOfName(i) > 5 Then
        lstscores.AddItem Array_of_players(i).Playername & vbTab & Array_of_players(i).score
    Else
        lstscores.AddItem Array_of_players(i).Playername & vbTab & vbTab & Array_of_players(i).score
    End If
Next i

End Sub

Private Sub txtsearch_Click()
txtsearch.Text = ""
End Sub
