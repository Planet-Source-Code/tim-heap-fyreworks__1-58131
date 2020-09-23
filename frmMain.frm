VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Fyreworks"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   499
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl Explosion 
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1455
      Left            =   2400
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnRunning As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


    'Particles are the little bits of shrapnel flying round the screen
Private Type Particles
    X As Integer
    Y As Integer
    XSpeed As Integer
    YSpeed As Integer
    Colour As Long
End Type

    'the colour of the firework
Private Type Colours
    r As Integer
    g As Integer
    b As Integer
End Type
    
    'the shrapnel with the tails
Private Type Sparks
    X As Integer
    Y As Integer
    OldX(1 To 5) As Integer 'the oldx and y positions of the spark
    OldY(1 To 5) As Integer 'used to make a trail behind them
    XSpeed As Integer
    YSpeed As Integer
    Colour As Colours
    Particle As Particles   'each spark has a particle assigned to it
End Type

    'the whole fyrework put together
Private Type FireWorks
    Spark(0 To 20) As Sparks
    Exploded As Boolean 'if the fyrework has exploded
    ExplosionTimer As Integer   'time to BOOM
    X As Integer
    Y As Integer
    XSpeed As Integer
    YSpeed As Integer
    SoundTimer As Integer       'time to the sound plays. this gives it a
                                'time delay as if it were far away
    Particle(1 To 10) As Particles  'Particle trail for the fyrework
    Colour As Long
End Type
    
    'all ten of the fireworks put together
Private FireWork(0 To 10) As FireWorks

Dim i As Integer, a As Integer, b As Integer    'for...next loops
Dim MouseDownX As Integer, MouseDownY As Integer    'used in the click event
Dim MouseX As Integer, MouseY As Integer            'makes the launching of the
Dim MouseIsDown As Boolean                          'fyreworks work
Dim TextTimer As Integer        'Time that the text is displayed

Const Message1 As String = "Fyreworks"  'text that is displayed
Const Message2 As String = "by Tim Heap"
Const ExplosionDelay As Integer = 20    'how long the fyrework wait to explode

Private Sub Form_KeyPress(KeyAscii As Integer)
blnRunning = False  'stops the loop
End Sub

Private Sub Form_Load()

On_Load 'does the preload stuff

Reset   'resets thee fyreworks

Main_Loop   'starts the loop up

    'deletes the unused sound files and closes the sound clip
For i = 0 To 10
    Explosion(i).Command = "close"
    Kill App.Path & "\fireworks" & i & ".wav"   'be careful with this statment
Next i

End 'end the program

End Sub

Sub On_Load()
Randomize   'randomizes stuff

    'resizing the form
Me.Height = Screen.Height
Me.Width = Screen.Width
Me.ScaleWidth = Me.Width / 15
Me.ScaleHeight = Me.Height / 15
Me.Top = 0
Me.Left = 0
Me.WindowState = 2
    'resizes the back buffer
picBackBuffer.Width = Me.ScaleWidth
picBackBuffer.Height = Me.ScaleHeight

    'resets variables
TextTimer = 100
blnRunning = True

    'loads the background
Picture1.Picture = LoadPicture(App.Path & "/CityScape.bmp")

    'Loads the sound file
    'this is from the MSDN help file: 'Multimedia MCI'
For i = 0 To 10
    If i <> 0 Then Load Explosion(i)
    Explosion(i).Notify = False
    Explosion(i).Wait = True
    Explosion(i).Shareable = False
    Explosion(i).DeviceType = "WaveAudio"
        'creates a new sound file
        'more than one mci control cant use the same file, so we make new ones
    FileCopy App.Path & "\fireworks.wav", App.Path & "\fireworks" & i & ".wav"
    Explosion(i).FileName = App.Path & "\fireworks" & i & ".wav"
    Explosion(i).Command = "Open"
Next i

Me.Show 'Shows the form
    'this is required because we never leave the form load event
End Sub

    'resets the fireworks
Sub Reset()
For i = 0 To 10
    With FireWork(i)
            'makes them all explode in a line
            'looks really cool
        .X = picBackBuffer.ScaleWidth / 11 * i + (picBackBuffer.ScaleWidth / 22)
        .Y = picBackBuffer.ScaleHeight / 4
    End With
Next i


End Sub

    'this is where every thing goes on
    'and i mean EVERYTHING
Sub Main_Loop()
    'sets up the loop
Dim lngTimeOfLastRender As Long
lngTimeOfLastRender = GetTickCount
Do While blnRunning = True  'loop while the loop is running ???
    If lngTimeOfLastRender + 20 < GetTickCount Then 'if weve waited 20 ms
        lngTimeOfLastRender = GetTickCount  'reset counter
        Move_Stuff  'move stuff?
        Draw_Stuff  'ummm draw stuff?
    End If
    DoEvents    'lets everything else have a go
Loop
End Sub

    'this sub is kinda stupid hey?
Sub Move_Stuff()
Move_FireWorks
End Sub

Sub Move_FireWorks()
For i = 0 To 10 'loops through all the fireworks
        'if its gone BOOM
    If FireWork(i).Exploded Then
            
            'checks to see if the sound needs playing
        If FireWork(i).SoundTimer <> -1 Then FireWork(i).SoundTimer = FireWork(i).SoundTimer - 1
        If FireWork(i).SoundTimer = 0 Then
            FireWork(i).SoundTimer = -1
            Explosion(i).Command = "prev"
            Explosion(i).Command = "play"
        End If
        
            'loops throguht the 'sparks'
        For a = 1 To 20
            With FireWork(i).Spark(a)
                    'moves the tail
                For b = 5 To 2 Step -1
                    .OldX(b) = .OldX(b - 1)
                    .OldY(b) = .OldY(b - 1)
                Next b
                .OldX(1) = .X
                .OldY(1) = .Y
                    
                    'moves the spark if its still on screen
                If .Y < Me.ScaleHeight Then
                    .YSpeed = .YSpeed + 1
                    If .X < 0 Then .XSpeed = Abs(.XSpeed)
                    If .X > Me.ScaleWidth Then .XSpeed = -Abs(.XSpeed)
                
                    .X = .X + .XSpeed
                    .Y = .Y + .YSpeed
                End If
                    
                    'dims the colour
                .Colour.r = .Colour.r / 1.1
                .Colour.g = .Colour.g / 1.1
                .Colour.b = .Colour.b / 1.1
                
                    'moves the particle asigned to the spark
                If .Particle.Y < Me.ScaleHeight Then
                    .Particle.YSpeed = .Particle.YSpeed + 1
                    .Particle.Y = .Particle.Y + .Particle.YSpeed
                    .Particle.X = .Particle.X + .Particle.XSpeed
                End If
            End With
        Next a
    Else    'if we havnt gone BOOM
        FireWork(i).ExplosionTimer = FireWork(i).ExplosionTimer - 1 'count down
            'moves the particles in its tail
        For a = LBound(FireWork(i).Particle) To UBound(FireWork(i).Particle)
                'launches the paritcle every so often
            If FireWork(i).ExplosionTimer = Int(ExplosionDelay / UBound(FireWork(i).Particle)) * a Then
                FireWork(i).Particle(a).X = FireWork(i).X
                FireWork(i).Particle(a).Y = FireWork(i).Y
                FireWork(i).Particle(a).XSpeed = FireWork(i).XSpeed + ((Rnd * 10) - 5)
                FireWork(i).Particle(a).YSpeed = FireWork(i).YSpeed + (Rnd * 10)
                FireWork(i).Particle(a).Colour = RGB(Rnd * 100, Rnd * 100, Rnd * 100)
            End If
        Next a
        If FireWork(i).ExplosionTimer <= 0 Then
                
            FireWork(i).Exploded = True
            FireWork(i).SoundTimer = 10 + (Rnd * 10)
            Dim intColour As Integer
            intColour = Int(3 * Rnd + 1)
            For a = 1 To 20
                With FireWork(i).Spark(a)
                    .X = FireWork(i).X
                    .Y = FireWork(i).Y
                    
                    For b = 1 To 5
                        .OldX(b) = .X
                        .OldY(b) = .Y
                    Next b
                    Dim intTemp As Integer
                    intTemp = Rnd * 20
                    .XSpeed = Sin(intTemp) * ((a / 2) + 5)
                    .YSpeed = Cos(intTemp) * ((a / 2) + 5)
                    Select Case intColour
                        Case 1
                            FireWork(i).Colour = RGB(255, 100, 100)
                            .Colour.r = (Rnd * 100) + 155
                            .Colour.g = Rnd * 100
                            .Colour.b = Rnd * 100
                        Case 2
                            FireWork(i).Colour = RGB(100, 255, 100)
                            .Colour.g = (Rnd * 100) + 155
                            .Colour.r = Rnd * 100
                            .Colour.b = Rnd * 100
                        Case 3
                            FireWork(i).Colour = RGB(100, 100, 255)
                            .Colour.b = (Rnd * 100) + 155
                            .Colour.g = Rnd * 100
                            .Colour.r = Rnd * 100
                    End Select
                    
                    .Particle.X = .X
                    .Particle.Y = .Y
                    .Particle.XSpeed = .XSpeed + ((Rnd * 10) - 5)
                    .Particle.YSpeed = .YSpeed + ((Rnd * 7) - 5)
                    .Particle.Colour = RGB(.Colour.r, .Colour.g, .Colour.b)
                    
                End With
            Next a
        Else
            FireWork(i).YSpeed = FireWork(i).YSpeed + 1
            FireWork(i).X = FireWork(i).X + FireWork(i).XSpeed
            FireWork(i).Y = FireWork(i).Y + FireWork(i).YSpeed
        End If
    
    End If
        'this moves the particle trail of the firework
        'if i put it in the loop above, it wouldn't move when the fyrework went BOOM
    For a = 1 To 10
        If FireWork(i).Particle(a).Y < Me.ScaleHeight Then
            FireWork(i).Particle(a).YSpeed = FireWork(i).Particle(a).YSpeed + 1
            FireWork(i).Particle(a).Y = FireWork(i).Particle(a).Y + FireWork(i).Particle(a).YSpeed
            FireWork(i).Particle(a).X = FireWork(i).Particle(a).X + FireWork(i).Particle(a).XSpeed
        End If
    Next a
Next i
End Sub

Sub Draw_Stuff()
    'declares stuff
Dim r1 As Integer, r2 As Integer
Dim g1 As Integer, g2 As Integer
Dim b1 As Integer, b2 As Integer
    'sets the colours
r1 = 0: g1 = 0: b1 = 0
r2 = 50: g2 = 0: b2 = 50
'colour 1 = black
'colour 2 = dark purple

    'will draw on the background
    'gradient from colour 1 to colour 2
For i = 0 To Me.ScaleHeight
    Dim r As Integer, g As Integer, b As Integer
    r = r1 + (((r2 - r1) / Me.ScaleHeight) * i) 'works out the colour we need
    g = g1 + (((g2 - g1) / Me.ScaleHeight) * i)
    b = b1 + (((b2 - b1) / Me.ScaleHeight) * i)
    picBackBuffer.Line (0, i)-(Me.ScaleWidth, i), RGB(r, g, b)  'draws it on
Next i

Draw_FireWorks  'draw the fyreworks on

    'Paint the city
BitBlt picBackBuffer.hDC, 0, Me.ScaleHeight - Picture1.ScaleHeight, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, vbSrcAnd

    'Draws the line if we have cliked and are ready to launch
If MouseIsDown Then
    picBackBuffer.DrawWidth = 3
    picBackBuffer.Line (MouseX, MouseY)-(MouseDownX, MouseDownY), RGB(255, 255, 255)
    picBackBuffer.DrawWidth = 1
End If

    'draws on the title at the start
If TextTimer > 0 Then
    TextTimer = TextTimer - 1
    picBackBuffer.FontSize = 20
    picBackBuffer.ForeColor = RGB((TextTimer / 100) * 255, (TextTimer / 100) * 255, (TextTimer / 100) * 255)
    picBackBuffer.CurrentX = (picBackBuffer.ScaleWidth / 2) - (picBackBuffer.TextWidth(Message1) / 2)
    picBackBuffer.CurrentY = picBackBuffer.ScaleHeight / 2 - picBackBuffer.TextHeight(Message1)
    picBackBuffer.Print Message1
    picBackBuffer.CurrentX = (picBackBuffer.ScaleWidth / 2) - (picBackBuffer.TextWidth(Message2) / 2)
    picBackBuffer.CurrentY = picBackBuffer.ScaleHeight / 2
    picBackBuffer.Print Message2
End If

    'sets the forms picture to the backbuffers picture
    'use this to stop filckering
Me.Picture = picBackBuffer.Image
End Sub

    'draws the fyreworks
Sub Draw_FireWorks()
For i = 0 To 10
    If FireWork(i).Exploded Then    'if its exploded
            'draws the explosion in the middle
            'i shouldn't really use the sound timer, but it was conveniant
        If FireWork(i).SoundTimer > 0 Then
                'draw a dot in the middle of the explosion
            picBackBuffer.DrawWidth = FireWork(i).SoundTimer * 2
            picBackBuffer.PSet (FireWork(i).X, FireWork(i).Y), FireWork(i).Colour
            picBackBuffer.DrawWidth = 1
        End If
        For a = 1 To 20
            With FireWork(i).Spark(a)
                    'draws the spark with the tail
                picBackBuffer.DrawWidth = 5
                picBackBuffer.Line (.X, .Y)-(.OldX(1), .OldY(1)), RGB(.Colour.r, .Colour.g, .Colour.b)
                For b = 1 To 4
                    picBackBuffer.DrawWidth = picBackBuffer.DrawWidth - 1
                    picBackBuffer.Line (.OldX(b), .OldY(b))-(.OldX(b + 1), .OldY(b + 1)), RGB((.Colour.r / 6) * (6 - b), (.Colour.g / 6) * (6 - b), (.Colour.b / 6) * (6 - b))
                Next b
                    'draws the particle
                picBackBuffer.DrawWidth = 3
                picBackBuffer.PSet (.Particle.X, .Particle.Y), .Particle.Colour
                picBackBuffer.DrawWidth = 1

            End With
        Next a
    Else    'draws the white firework dot
        picBackBuffer.DrawWidth = 10
        Dim intTemp As Integer
            'this makes it blak at first, going to white when its about to explode
        intTemp = ((ExplosionDelay - FireWork(i).ExplosionTimer) / ExplosionDelay) * 255
        picBackBuffer.PSet (FireWork(i).X, FireWork(i).Y), RGB(intTemp, intTemp, intTemp)
        picBackBuffer.DrawWidth = 1
    End If
        'draws the particle trail
    For a = 1 To 10
        With FireWork(i).Particle(a)
            picBackBuffer.DrawWidth = 3
            picBackBuffer.PSet (.X, .Y), .Colour
            picBackBuffer.DrawWidth = 1
        End With
    Next a
Next i
End Sub

    'this bit launches the fyreworks
    'pretty easy to follow
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDownX = X
MouseDownY = Y
MouseIsDown = True
MouseX = X
MouseY = Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X
MouseY = Y
End Sub

    'launch the fyrework
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static FireWorkNumber As Integer
FireWorkNumber = FireWorkNumber + 1
If FireWorkNumber > 10 Then FireWorkNumber = 1
i = FireWorkNumber
    'resets variables
FireWork(i).Exploded = False
FireWork(i).ExplosionTimer = ExplosionDelay
    'sets its position and speed
FireWork(i).X = X
FireWork(i).Y = Y
FireWork(i).XSpeed = (MouseDownX - X) / 10
FireWork(i).YSpeed = (MouseDownY - Y) / 10

MouseIsDown = False
End Sub

Private Sub Form_Terminate()
    'stops the loop running
blnRunning = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
blnRunning = False
End Sub
