VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Magnet"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2520
   LinkTopic       =   "Form13"
   ScaleHeight     =   3090
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox List1 
      Height          =   315
      ItemData        =   "Form13.frx":0000
      Left            =   1080
      List            =   "Form13.frx":000A
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Fuerza 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Y 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox X 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Activated:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Force:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   135
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Magneton.X = Val(X.Text)
Magneton.Y = Val(Y.Text)
Magneton.Force = Val(Fuerza.Text)

If List1.Text = "Si" Then
    Magneton.Activated = True
Else
    Magneton.Activated = False
End If


End Sub

