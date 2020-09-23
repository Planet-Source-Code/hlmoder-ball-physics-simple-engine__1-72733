Attribute VB_Name = "Module1"
'=======================================================
'Simple Ball Physics Engine Programed By J.Rafael Cenit (also knowed as HLMODER)
'If you use any part of this code please give me credits
'=======================================================

Public Type Colors 'Colors for the Ball
    R As Integer
    G As Integer
    B As Integer
End Type

Public Type ESFERA 'The Ball it selfs
    X As Single
    Y As Single
    vX As Double
    vY As Double
    Weigth As Double
    Elasticity As Double
    Material As String
    Color As Colors
End Type


Public Type magnet 'The Magnet
    X As Single
    Y As Single
    Force As Single
    Activated As Boolean
End Type

Public Dist As Single 'Distance
Public Gravity As Single 'Gravity
Public Wind As Single 'Wind
Public Ball(0 To 40) As ESFERA 'Ball Array
Public nBalls As Integer 'Numbers of Balls
Public Magneton As magnet 'The Magnet
Public CurX, CurY As Integer 'Current X,Y mouse coordinates
Public i As Integer

'Insert a New Ball
Public Sub InsertNewBall()
If nBalls < 40 Then
    nBalls = nBalls + 1
    
    Ball(nBalls - 1).Weigth = Rnd * 40
    Ball(nBalls - 1).Color.R = 10 + Int(Rnd * 255)
    Ball(nBalls - 1).Color.G = 2 + Int(Rnd * 255)
    Ball(nBalls - 1).Color.B = 3 + Int(Rnd * 255)
    Ball(nBalls - 1).X = CurX
    Ball(nBalls - 1).Y = CurY
    Ball(nBalls - 1).vX = Int(Rnd * 200)
    Ball(nBalls - 1).vY = 0
    Ball(nBalls - 1).Elasticity = Rnd * 5
    Ball(nBalls - 1).Material = ""

    
Else
    MsgBox "Limit Reached"
End If

End Sub

'Here Comes The Interesting Part... Engine It Selfs
Public Sub PhysicsEngine()

Form10.Picture1.Cls 'Clear The Frame

For i = 0 To nBalls - 1


     
            'Check Y Limits
            If Magneton.Activated = True Then 'Is The Magnet activated?
            
                    If Ball(i).Material = "Metal" Then 'Is The Current Ball a Metal Ball?
        
                            'If Yes Check The Distance Between The Current Ball and The Magnet
                            Dist = ((Magneton.X - Ball(i).X) * (Magneton.X - Ball(i).X) + (Magneton.Y - Ball(i).Y) * (Magneton.Y - Ball(i).Y)) ^ 0.5
                               
                               If Dist <= 260 Then 'Dont Let The Ball Go Inside Magnet
                                   Ball(i).X = Magneton.X - 260
                               Else
                                   Ball(i).vY = (Magneton.Y - Ball(i).Y) / Dist
                                   Ball(i).Y = Ball(i).Y + (Ball(i).vY * Magneton.Force)
                                   
                                   Ball(i).vX = (Magneton.X - Ball(i).X) / Dist
                                   Ball(i).X = Ball(i).X + (Ball(i).vX * Magneton.Force)
                               End If
                               
                        End If
                        
                Picture1.Circle (Magneton.X, Magneton.Y), 200, vbWhite 'Draw Magnet
                
            End If

  
            If Ball(i).Y >= Form10.Picture1.Height - 160 Then 'If The Ball Is "on the floor" Then Bounce
                Ball(i).Y = Form10.Picture1.Height - 160
                Ball(i).vY = (Ball(i).vY + (-Ball(i).Weigth + Ball(i).Elasticity)) * -0.6
            End If
            
            If Ball(i).Y < 160 Then 'If The Ball Is "on the roof" Then Bounce
                Ball(i).Y = 160
                Ball(i).vY = (Ball(i).vY + (-Ball(i).Weigth + Ball(i).Elasticity)) * -0.6
            End If
        
        
        
             'Check X limits
             If Ball(i).X >= Form10.Picture1.Width - 140 Then 'If The Ball Is "on the right wall" Then Bounce
                 Ball(i).X = Form10.Picture1.Width - 140
                 Ball(i).vX = (Ball(i).vX + (-Ball(i).Weigth + Ball(i).Elasticity)) * -0.6
             End If
             
             If Ball(i).X < 140 Then 'If The Ball Is "on the left wall" Then Bounce
                 Ball(i).X = 140
                 Ball(i).vX = (Ball(i).vX + (-Ball(i).Weigth + Ball(i).Elasticity)) * -0.6
             End If
            
    Ball(i).vY = Ball(i).vY - Gravity * (Ball(i).Weigth / 20) 'Apply Y Variation
    Ball(i).Y = Ball(i).Y + Ball(i).vY 'Update Y Position
    Ball(i).vX = Ball(i).vX + Wind  'Apply  Variation
    Ball(i).X = Ball(i).X + Ball(i).vX 'Update X Position

    Form10.Picture1.Circle (Ball(i).X, Ball(i).Y), 100, RGB(Ball(i).Color.R, Ball(i).Color.G, Ball(i).Color.B) 'Draw Ball
  
Next i

End Sub
