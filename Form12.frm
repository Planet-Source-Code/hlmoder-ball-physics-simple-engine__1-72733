VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "New Ball"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2835
   LinkTopic       =   "Form12"
   ScaleHeight     =   3090
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox List1 
      Height          =   315
      ItemData        =   "Form12.frx":0000
      Left            =   1080
      List            =   "Form12.frx":000A
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Elasticity 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Weigth 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Material:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Elasticity:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Weigth:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If nBalls < 40 Then
    nBalls = nBalls + 1
    
    Ball(nBalls - 1).Weigth = Val(Weigth.Text)
    Ball(nBalls - 1).Color.R = 10 + Int(Rnd * 255)
    Ball(nBalls - 1).Color.G = 2 + Int(Rnd * 255)
    Ball(nBalls - 1).Color.B = 3 + Int(Rnd * 255)
    Ball(nBalls - 1).X = Rnd * 5000
    Ball(nBalls - 1).Y = 500
    Ball(nBalls - 1).vX = Int(Rnd * 200)
    Ball(nBalls - 1).vY = 0
    Ball(nBalls - 1).Elasticity = Val(Elasticity.Text)
    Ball(nBalls - 1).Material = List1.Text
    
Else
    MsgBox "Limit Reached"
End If

End Sub



