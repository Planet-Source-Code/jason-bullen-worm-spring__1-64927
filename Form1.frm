VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Click in window to move WORM"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   446
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   607
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1200
      Top             =   1800
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "LARGE 2"
      Height          =   495
      Index           =   4
      Left            =   5400
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "LARGE"
      Height          =   495
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "PAUSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "BASIC"
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "DOUBLE"
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "SINGLE"
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      Height          =   5655
      Left            =   0
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   565
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   3720
         Top             =   3480
      End
      Begin VB.Label lblFPS 
         Caption         =   "FPS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0



Private Const SPRING_STIFF As Single = 0.2
Private Const SPRING_DAMP As Single = 0.2

Private mFrames As Long
Private mGo As Boolean


Private colParticles As New Collection
Private colSprings As New Collection


Private Sub VecCopy(vec1 As clsVector, vecRes As clsVector)
    vecRes.x = vec1.x
    vecRes.y = vec1.y
End Sub

Private Sub VecAdd(vec1 As clsVector, vec2 As clsVector, vecRes As clsVector)
    vecRes.x = vec1.x + vec2.x
    vecRes.y = vec1.y + vec2.y
End Sub

Private Sub VecSub(vec1 As clsVector, vec2 As clsVector, vecRes As clsVector)
    vecRes.x = vec1.x - vec2.x
    vecRes.y = vec1.y - vec2.y
End Sub

Private Sub VecScale(vec1 As clsVector, s As Single, vecRes As clsVector)
    vecRes.x = vec1.x * s
    vecRes.y = vec1.y * s
End Sub

Private Function VecDot(vec1 As clsVector, vec2 As clsVector) As Single
    VecDot = vec1.x * vec2.x + vec1.y * vec2.y
End Function

Private Function VecLength(vec1 As clsVector) As Single
    VecLength = Sqr(vec1.x * vec1.x + vec1.y * vec1.y)
End Function


' This piece of code uses Mass, Force, and time to move a particle
Private Sub ParIntegrate(parSource As clsParticle, parDest As clsParticle, deltaTime As Single)
    Dim deltaOneOverMass As Single
    
    deltaOneOverMass = deltaTime * parSource.oneOverMass
    
    parDest.v.x = parSource.v.x + (parSource.f.x * deltaOneOverMass)
    parDest.v.y = parSource.v.y + (parSource.f.y * deltaOneOverMass)

    parDest.pos.x = parSource.pos.x + (deltaTime * parSource.v.x)
    parDest.pos.y = parSource.pos.y + (deltaTime * parSource.v.y)
End Sub

' This code accumulates forces applied by springs
Private Sub ParSpring(par1 As clsParticle, par2 As clsParticle, restLen As Single)
    Dim dist As Single
    Dim hTerm As Single
    Dim dTerm As Single
    Dim deltaP As New clsVector
    Dim deltaV As New clsVector
    Dim sprForce As New clsVector
    
    Call VecSub(par1.pos, par2.pos, deltaP)
    dist = VecLength(deltaP)

    hTerm = (dist - restLen) * SPRING_STIFF
    
    Call VecSub(par1.v, par2.v, deltaV)

    dTerm = (VecDot(deltaV, deltaP) * SPRING_DAMP) / dist
    
    Call VecScale(deltaP, 1# / dist, sprForce)
    
    Call VecScale(sprForce, -(hTerm + dTerm), sprForce)

    Call VecAdd(par1.f, sprForce, par1.f)
    Call VecSub(par2.f, sprForce, par2.f)
End Sub



Private Sub DrawStuff()
    Dim i As Integer
    Dim par1 As clsParticle
    Dim par2 As clsParticle
    Dim spr1 As clsSpring
    
    picMain.Cls
    
'    picMain.DrawWidth = 2
    
    For i = 1 To colParticles.Count
        Set par1 = colParticles.Item(i)
        picMain.Circle (par1.pos.x, par1.pos.y), par1.radius, par1.color
    Next
    
'    picMain.DrawWidth = 2
    
    For i = 1 To colSprings.Count
        Set spr1 = colSprings.Item(i)
        Set par1 = colParticles.Item(spr1.index1)
        Set par2 = colParticles.Item(spr1.index2)
       
        picMain.Line (par1.pos.x, par1.pos.y)-(par2.pos.x, par2.pos.y), spr1.color
    Next
End Sub


Private Sub ApplyForces()
    Dim i As Integer
    Dim par1 As clsParticle
    Dim par2 As clsParticle
    Dim spr1 As clsSpring

    ' Reset forces to Zero
    For i = 1 To colParticles.Count
        Set par1 = colParticles.Item(i)
        par1.f.x = 0#
        par1.f.y = 0#
    Next

    ' Accumulate forces by applying springs
    For i = 1 To colSprings.Count
        Set spr1 = colSprings.Item(i)
        Set par1 = colParticles.Item(spr1.index1)
        Set par2 = colParticles.Item(spr1.index2)
        Call ParSpring(par1, par2, spr1.restLen)
    Next

    ' Apply forces to particles and move them
    For i = 1 To colParticles.Count
        Set par1 = colParticles.Item(i)
        Call ParIntegrate(par1, par1, 1#)
    Next
End Sub


'BELOW here is very boring
'-------------------------
'-------------------------

Private Sub cmdStart_Click(Index As Integer)
    Dim par1 As clsParticle
    Dim spr1 As clsSpring
    Dim i  As Integer
    
    For i = 0 To 4
        cmdStart(i).Enabled = False
    Next

    ' All these buttons simply create Particles and Springs
    ' The same spring and integrator code is used by all

    Select Case Index
    Case 0
        Set par1 = New clsParticle
        par1.pos.x = 100
        par1.pos.y = 100
        par1.oneOverMass = 1# / 38#
        par1.radius = 180
        par1.color = Rnd() * &HFFFFFF
        Call colParticles.Add(par1)
        
        ' Create particles
        For i = 2 To 100
            Set par1 = New clsParticle
            par1.pos.x = 100 + i * 4
            par1.pos.y = 100
            par1.oneOverMass = 1# / 2#
            par1.radius = 20
            par1.color = Rnd() * &HFFFFFF
            Call colParticles.Add(par1)
        Next
        
        ' Create main springs
        For i = 1 To 99
            Set spr1 = New clsSpring
            spr1.index1 = i
            spr1.index2 = i + 1
            spr1.restLen = 4
            spr1.color = RGB(255, 0, 0)
            Call colSprings.Add(spr1)
        Next
    
    Case 1
        ' Create particles
        For i = 1 To 100
            Set par1 = New clsParticle
            par1.pos.x = 100 + i * 4
            par1.pos.y = 100
            par1.oneOverMass = 1# / 1#
            par1.radius = 10
            par1.color = Rnd() * &HFFFFFF
            Call colParticles.Add(par1)
        Next
        
        ' Create main springs
        For i = 1 To 99
            Set spr1 = New clsSpring
            spr1.index1 = i
            spr1.index2 = i + 1
            spr1.restLen = 4
            spr1.color = RGB(255, 0, 0)
            Call colSprings.Add(spr1)
        Next
    
        ' Create springs to straighten worm
        For i = 1 To 98
            Set spr1 = New clsSpring
            spr1.index1 = i
            spr1.index2 = i + 2
            spr1.restLen = 8
            spr1.color = RGB(0, 0, 255)
            Call colSprings.Add(spr1)
        Next
    
    Case 2
        ' Create particles
        Set par1 = New clsParticle
        par1.pos.x = 100
        par1.pos.y = 100
        par1.oneOverMass = 1# / 1#
        par1.radius = 10
        par1.color = Rnd() * &HFFFFFF
        Call colParticles.Add(par1)
        
        Set par1 = New clsParticle
        par1.pos.x = 200
        par1.pos.y = 150
        par1.oneOverMass = 1# / 1.5
        par1.radius = 15
        par1.color = Rnd() * &HFFFFFF
        Call colParticles.Add(par1)
        
        Set par1 = New clsParticle
        par1.pos.x = 300
        par1.pos.y = 100
        par1.oneOverMass = 1# / 30#
        par1.radius = 30
        par1.color = Rnd() * &HFFFFFF
        Call colParticles.Add(par1)
        
        ' Create springs to straighten worm
        Set spr1 = New clsSpring
        spr1.index1 = 1
        spr1.index2 = 2
        spr1.restLen = 150
        Call colSprings.Add(spr1)
        
        Set spr1 = New clsSpring
        spr1.index1 = 2
        spr1.index2 = 3
        spr1.restLen = 50
        Call colSprings.Add(spr1)
    
    Case 3
        ' Create particles
        For i = 1 To 20
            Set par1 = New clsParticle
            par1.pos.x = 100 + i * 20
            par1.pos.y = 100
            par1.oneOverMass = 1# / 3#
            par1.radius = 30
            par1.color = Rnd() * &HFFFFFF
            Call colParticles.Add(par1)
        Next
        
        ' Create main springs
        For i = 1 To 19
            Set spr1 = New clsSpring
            spr1.index1 = i
            spr1.index2 = i + 1
            spr1.restLen = 20
            spr1.color = RGB(255, 0, 0)
            Call colSprings.Add(spr1)
        Next
    
    Case 4
        ' Create particles
        For i = 1 To 20
            Set par1 = New clsParticle
            par1.pos.x = 100 + i * 20
            par1.pos.y = 100
            par1.oneOverMass = 1# / 4#
            par1.radius = 40
            par1.color = Rnd() * &HFFFFFF
            Call colParticles.Add(par1)
        Next
        
        ' Create main springs
        For i = 1 To 19
            Set spr1 = New clsSpring
            spr1.index1 = i
            spr1.index2 = i + 1
            spr1.restLen = 20
            spr1.color = RGB(255, 0, 0)
            Call colSprings.Add(spr1)
        Next
    
            ' Create springs to straighten worm
        For i = 1 To 18
            Set spr1 = New clsSpring
            spr1.index1 = i
            spr1.index2 = i + 2
            spr1.restLen = 40
            spr1.color = RGB(0, 0, 255)
            Call colSprings.Add(spr1)
        Next
    End Select
    
    mGo = True
End Sub




Private Sub Form_Resize()
    Dim i As Integer
    If Form1.Width < 9600 Then Form1.Width = 9600
    If Form1.Height < 7200 Then Form1.Height = 7200

    For i = 0 To 4
        cmdStart(i).Top = ScaleHeight - cmdStart(i).Height - 8
    Next
    
    cmdPause.Top = ScaleHeight - cmdPause.Height - 8
    
    picMain.Height = ScaleHeight - cmdStart(0).Height - 16
    picMain.Width = ScaleWidth
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim par1 As clsParticle
    
    If colParticles.Count = 0 Then Exit Sub
    
    If Button <> 0 Then
        Set par1 = colParticles.Item(1)
        par1.pos.x = x
        par1.pos.y = y
    End If
End Sub



Private Sub Timer1_Timer()
    If mGo Then
        mFrames = mFrames + 1
        Call ApplyForces
        Call DrawStuff
    End If
End Sub

Private Sub Timer2_Timer()
    lblFPS.Caption = "FPS:" & Str(mFrames)
    mFrames = 0
End Sub

Private Sub cmdPause_Click()
    mGo = Not mGo
End Sub
