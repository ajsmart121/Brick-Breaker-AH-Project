VERSION 5.00
Begin VB.Form frminstructions 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instructions"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9330
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstinstructions 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   4890
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
   Begin VB.Label lblstart 
      BackColor       =   &H00000000&
      Caption         =   "Start Game!"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblquit 
      BackColor       =   &H00000000&
      Caption         =   "Quit To Menu"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frminstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim instructions(15) As String

FileName = App.Path & "\instructions.txt"

Open FileName For Input As #1
    For i = 0 To 13
        Input #1, instructions(i)
        lstinstructions.AddItem instructions(i)
    Next i
Close #1

'** read in the instructions file and display in a list box

End Sub

Private Sub lblquit_Click()
frmhome.Enabled = True
frmhome.Show

frminstructions.Hide
frminstructions.Enabled = False

'** if the quit button is pressed, then show the home form and hide the instructions form

End Sub

Private Sub lblstart_Click()
frmgame.Enabled = True
frmgame.Show

frminstructions.Hide
frminstructions.Enabled = False

'** if the start button is pressed, then show the game form and hide the instructions form

End Sub
