VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Craps v1.0"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4335
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame fraPoint 
      BackColor       =   &H00008000&
      Caption         =   "Point"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.Image imgPointDie2 
         Height          =   735
         Left            =   1320
         Top             =   360
         Width           =   855
      End
      Begin VB.Image imgPointDie1 
         Height          =   735
         Left            =   240
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image imgDie2 
      Height          =   735
      Left            =   2160
      Top             =   2520
      Width           =   855
   End
   Begin VB.Image imgDie1 
      Height          =   735
      Left            =   360
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMyPoint As Integer
Dim mDie1 As Integer
Dim mDie2 As Integer

Enum DiceNames
    SnakeEyes = 2
    trey
    [yo leven] = 11
    boxcars
End Enum
Private Function RollDice() As Integer
    Dim die1 As Integer, die2 As Integer, diesum As Integer
    Dim a As Integer, b As Integer
    
    die1 = 1 + Int(6 * Rnd()) 'Roll die1
    die2 = 1 + Int(6 * Rnd()) 'Roll die2
    
    Call DisplayDie(imgDie1, die1)  'Draw die1 image
    Call DisplayDie(imgDie2, die2)  'Draw die2 image
    
    mDie1 = die1            'Store die1 value
    mDie2 = die2            'Store die2 value
    diesum = die1 + die2    'Sum dice
    RollDice = diesum       'Return dieSum to caller
    
End Function

Private Sub DisplayDie(imgDie As Image, face As Integer)
    imgDie.Picture = LoadPicture(App.Path & "\images\die" & face & ".gif")
End Sub

    



Private Sub cmdPlay_Click()
Dim sum As Integer

mMyPoint = 0
fraPoint.Caption = "Point"
lblStatus.Caption = ""
imgPointDie1.Picture = LoadPicture("")
imgPointDie2.Picture = LoadPicture("")
Call Randomize

sum = RollDice()

Select Case sum
    Case 7, [yo leven]
        cmdRoll.Enabled = False
        lblStatus.Caption = "You Win!!!"
    Case SnakeEyes, trey, boxcars
        cmdRoll.Enabled = False
        lblStatus.Caption = "Sorry, you lose."
    Case Else
        mMyPoint = sum
        fraPoint.Caption = "Point is " & sum
        lblStatus.Caption = "Roll Again."
        Call DisplayDie(imgPointDie1, mDie1)
        Call DisplayDie(imgPointDie2, mDie2)
        cmdPlay.Enabled = False
        cmdRoll.Enabled = True
    End Select
    
        
End Sub

Private Sub cmdRoll_Click()
Dim sum As Integer

sum = RollDice()

If sum = mMyPoint Then
    lblStatus.Caption = "You Win"
    cmdRoll.Enabled = False
    cmdPlay.Enabled = True
ElseIf sum = 7 Then
    lblStatus.Caption = "Sorry. You Lose."
    cmdRoll.Enabled = False
    cmdPlay.Enabled = True
End If

End Sub

Private Sub Form_Load()
Icon = LoadPicture(App.Path & "\images\die.ICO")

End Sub
