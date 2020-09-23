VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ball Physics Engine"
   ClientHeight    =   9600
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   12840
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Punch 
      Caption         =   "Random Punch"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton WPlus 
      Caption         =   "+"
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton WMinus 
      Caption         =   "-"
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton GMinus 
      Caption         =   "-"
      Height          =   195
      Left            =   2160
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton GPlus 
      Caption         =   "+"
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox GravityVar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   4
      Text            =   "-10"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox WindVar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   9240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   12825
      TabIndex        =   0
      ToolTipText     =   "Click To Insert New Ball With Random Values"
      Top             =   1080
      Width           =   12855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Gravity"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wind"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu newball 
         Caption         =   "New Ball"
      End
      Begin VB.Menu magnet 
         Caption         =   "Magnet"
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
Randomize

nBalls = 1
Gravity = -10
Wind = 0

For i = 0 To nBalls - 1
    Ball(i).vX = 0
    Ball(i).vY = 0
    Ball(i).X = 500 + Int((Rnd * 4000))
    Ball(i).Y = 100 + Int((Rnd * 600))
    Ball(i).Color.R = 10 + Int(Rnd * 255)
    Ball(i).Color.G = 2 + Int(Rnd * 255)
    Ball(i).Color.B = 3 + Int(Rnd * 255)
    Ball(i).Weigth = 10
Next i

End Sub


Private Sub GMinus_Click()
GravityVar.Text = GravityVar.Text - 1
Gravity = Val(GravityVar.Text)
End Sub


Private Sub GPlus_Click()
GravityVar.Text = GravityVar.Text + 1
Gravity = Val(GravityVar.Text)
End Sub


Private Sub magnet_Click()
Form13.Show
End Sub

Private Sub newball_Click()
Form12.Show
End Sub

Private Sub Picture1_Click()
Call InsertNewBall
End Sub




Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CurX = X
CurY = Y
End Sub

Private Sub Punch_Click()
For i = 0 To nBalls - 1
    Ball(i).vY = Rnd * 500
    
    If nBalls Mod 2 = 0 Then
        Ball(i).vX = Rnd * 400
    Else
        Ball(i).vX = Rnd * -400
    End If
Next i
End Sub

Private Sub Timer1_Timer()
Call PhysicsEngine
End Sub

Private Sub WMinus_Click()
WindVar.Text = WindVar.Text - 1
Wind = Val(WindVar.Text)
End Sub

Private Sub WPlus_Click()
WindVar.Text = WindVar.Text + 1
Wind = Val(WindVar.Text)
End Sub
