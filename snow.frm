VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D Snow"
   ClientHeight    =   8265
   ClientLeft      =   2385
   ClientTop       =   1935
   ClientWidth     =   9585
   ControlBox      =   0   'False
   FillColor       =   &H80000003&
   Icon            =   "snow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   639
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3480
      Max             =   20
      Min             =   1
      TabIndex        =   19
      Top             =   7800
      Value           =   1
      Width           =   1695
   End
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Text            =   "3"
      Top             =   7800
      Width           =   615
   End
   Begin VB.CheckBox chkRotate 
      Caption         =   "Rotating Flakes"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   7560
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton cmdFullScreen 
      Caption         =   "&Full Screen"
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Timer tmrFlakeHorzRotate 
      Interval        =   1000
      Left            =   7680
      Top             =   7320
   End
   Begin VB.PictureBox picSnowLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   99
      Left            =   6360
      Picture         =   "snow.frx":014A
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   13
      Top             =   8640
      Width           =   870
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Close"
      Height          =   495
      Left            =   8160
      TabIndex        =   12
      ToolTipText     =   "Press ALT-C to Quit"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.CheckBox chkWrap 
      Caption         =   "Wrap-around Snow"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   8040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkRandSpeed 
      Caption         =   "Random Speed"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   7800
      Width           =   1455
   End
   Begin VB.CheckBox chkDots 
      Caption         =   "Big Dots"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   7320
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtDots 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "500"
      Top             =   7560
      Width           =   615
   End
   Begin VB.TextBox txtFlakes 
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Text            =   "50"
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply Changes"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   7800
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7200
      Left            =   15
      Picture         =   "snow.frx":157C
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   9615
   End
   Begin VB.PictureBox picSnowSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6480
      Picture         =   "snow.frx":FCFF
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox picSnowLarge 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   7320
      Picture         =   "snow.frx":102A5
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   58
      TabIndex        =   1
      Top             =   8640
      Width           =   870
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   476
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   637
      TabIndex        =   0
      Top             =   0
      Width           =   9615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Background:"
      Height          =   255
      Left            =   3480
      TabIndex        =   20
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Rotation Degr."
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label lblAngle 
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   8640
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Snow Dots"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Snow Flakes"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB Snow
'Portions from vb-world.net
'And from planet-source-code.com
'Majority of the code (either original or modified) -
'Zane Horton
'Background pictures from www.planetside.co.uk - Home of Terragen
'Thanks to the original author of Snowflakes on Planet Source Code
'From which this got started when I said, 'I wonder how fast this would
'be with Setpixel instead of Pset...'
'Set SetPixelIt to FALSE in the Sub cmdApply_Click()
'if you want to see how much of a difference it makes...

'Start array index with 1 not 0.
Option Base 1
Option Explicit

Dim FLAKEROTATEANGLE As Integer

Dim I As Integer
Dim j As Integer

Dim AvgFPS As Single
Dim TempDBL As Double
Dim NumFrames As Long

Dim FullScreenSize As Boolean
Dim OrgScreenX As Integer
Dim OrgScreenY As Integer
Dim OrgScreenBPP As Byte

Dim Max_Snow As Integer
Dim Max_Flakes As Integer
Dim BigFlakes As Boolean
Dim setPixelIt As Boolean
Dim SnowWrap As Boolean
Dim FullScreen As Boolean
Dim RandSpeed As Boolean
Dim RotatingFlakes As Boolean

Dim Backgrounds(99) As String

Private Type Snow
    X As Integer
    Y As Integer
    Z As Integer
    Speed As Integer
    Wind As Integer
    CLR As Long
End Type

Private Type SnowFlake
    X As Long
    Y As Long
    Speed As Integer
    Wind As Integer
    LargeFlake As Boolean
    LastX As Long
    lastY As Long
    FlakeNum As Byte
End Type

Dim Snow(9999) As Snow
Dim SnowFlakes(9999) As SnowFlake
Dim Ended As Boolean

Sub DoSnow() 'Create & Animate the snow.

    Dim tmr1 As New clsTimer
    Dim FPS As Single
    Dim NumBlts As Long
    Dim NumSets As Long
    Dim tempSNG As Single
    
    tmr1.StartTimer
    
    BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture2.hdc, 0, 0, SRCCOPY
    'Picture1.Cls
    
    For I = 1 To Max_Snow
        'Allow other system events. (So it doesn't freeze)
        DoEvents
            'Calculate New Position
            Snow(I).Y = Snow(I).Y + Snow(I).Speed
            If SnowWrap Then
                If Snow(I).X > Picture1.ScaleWidth Then Snow(I).X = 0
            End If
            Snow(I).X = Snow(I).X + Snow(I).Wind
            'Erase Old and Draw new
            If setPixelIt Then
                If BigFlakes Then
                    SetPixel Picture1.hdc, (Snow(I).X), (Snow(I).Y), Snow(I).CLR
                    SetPixel Picture1.hdc, (Snow(I).X), (Snow(I).Y) + 1, Snow(I).CLR
                    SetPixel Picture1.hdc, (Snow(I).X) + 1, (Snow(I).Y), Snow(I).CLR
                    SetPixel Picture1.hdc, (Snow(I).X) + 1, (Snow(I).Y) + 1, Snow(I).CLR
                  Else
                    SetPixel Picture1.hdc, (Snow(I).X), (Snow(I).Y), Snow(I).CLR
                End If
              Else
                If BigFlakes Then
                    Picture1.PSet ((Snow(I).X), (Snow(I).Y)), Snow(I).CLR
                    Picture1.PSet ((Snow(I).X), (Snow(I).Y) + 1), Snow(I).CLR
                    Picture1.PSet ((Snow(I).X) + 1, (Snow(I).Y)), Snow(I).CLR
                    Picture1.PSet ((Snow(I).X) + 1, (Snow(I).Y) + 1), Snow(I).CLR
                  Else
                    Picture1.PSet ((Snow(I).X), (Snow(I).Y)), Snow(I).CLR
                End If
            End If
            
            'If snow is offscreen, bring it back to the top
            'and give it a new X axis location.
            
            If Snow(I).Y > Picture1.ScaleHeight - 10 Then
                If GetPixel(Picture1.hdc, Snow(I).X + Snow(I).Wind, Snow(I).Y + Snow(I).Speed) = vbWhite Then
                    If setPixelIt Then
                        If BigFlakes Then
                            SetPixel Picture1.hdc, (Snow(I).X + Snow(I).Wind), (Snow(I).Y + Snow(I).Speed) - 1, Snow(I).CLR
                            SetPixel Picture1.hdc, (Snow(I).X + Snow(I).Wind) + 1, (Snow(I).Y + Snow(I).Speed) - 1, Snow(I).CLR
                            SetPixel Picture1.hdc, (Snow(I).X + Snow(I).Wind), (Snow(I).Y + Snow(I).Speed), Snow(I).CLR
                            SetPixel Picture1.hdc, (Snow(I).X + Snow(I).Wind) + 1, (Snow(I).Y + Snow(I).Speed), Snow(I).CLR
                          Else
                            SetPixel Picture1.hdc, (Snow(I).X + Snow(I).Wind), (Snow(I).Y + Snow(I).Speed) - 1, Snow(I).CLR
                        End If
                      Else
                        If BigFlakes Then
                            Picture1.PSet ((Snow(I).X + Snow(I).Wind), ((Snow(I).Y + Snow(I).Speed) - 1)), Snow(I).CLR
                            Picture1.PSet ((Snow(I).X + Snow(I).Wind), ((Snow(I).Y + Snow(I).Speed) - 1) + 1), Snow(I).CLR
                            Picture1.PSet ((Snow(I).X + Snow(I).Wind) + 1, ((Snow(I).Y + Snow(I).Speed) - 1)), Snow(I).CLR
                            Picture1.PSet ((Snow(I).X + Snow(I).Wind) + 1, ((Snow(I).Y + Snow(I).Speed) - 1) + 1), Snow(I).CLR
                          Else
                            Picture1.PSet (Snow(I).X + Snow(I).Wind, (Snow(I).Y + Snow(I).Speed) - 1), Snow(I).CLR
                        End If
                    
                    End If
                    
                End If

                If Snow(I).Y > Picture1.ScaleHeight - 10 Then
                    If setPixelIt Then
                        If BigFlakes Then
                            SetPixel Picture1.hdc, (Snow(I).X), (Snow(I).Y), Snow(I).CLR
                            SetPixel Picture1.hdc, (Snow(I).X), (Snow(I).Y) + 1, Snow(I).CLR
                            SetPixel Picture1.hdc, (Snow(I).X) + 1, (Snow(I).Y), Snow(I).CLR
                            SetPixel Picture1.hdc, (Snow(I).X) + 1, (Snow(I).Y) + 1, Snow(I).CLR
                          Else
                            SetPixel Picture1.hdc, (Snow(I).X), (Snow(I).Y), Snow(I).CLR
                        End If
                      Else
                        If BigFlakes Then
                            Picture1.PSet (Snow(I).X, Snow(I).Y), Snow(I).CLR
                            Picture1.PSet (Snow(I).X, Snow(I).Y + 1), Snow(I).CLR
                            Picture1.PSet (Snow(I).X + 1, Snow(I).Y), Snow(I).CLR
                            Picture1.PSet (Snow(I).X + 1, Snow(I).Y + 1), Snow(I).CLR
                          Else
                            Picture1.PSet (Snow(I).X, Snow(I).Y), Snow(I).CLR
                        End If
                    End If
                    Snow(I).X = Int(Rnd * Picture1.ScaleWidth) - 35
                    Snow(I).Y = 0
                End If


            End If
            'Check if Ended.
            If Ended = True Then GoTo E
        Next I

    'Process the flakes
    For j = 1 To Max_Flakes
        If Ended = True Then GoTo E
        SnowFlakes(j).Y = SnowFlakes(j).Y + SnowFlakes(j).Speed
        If SnowWrap Then
            If SnowFlakes(j).X > Picture1.ScaleWidth Then SnowFlakes(j).X = 0
            If SnowFlakes(j).Y > Picture1.ScaleHeight Then SnowFlakes(j).Y = 0
        End If
        SnowFlakes(j).X = SnowFlakes(j).X + SnowFlakes(j).Wind
        
        'Erase the old Flakes
        If SnowFlakes(j).LastX <> 99999 Then
            'Do nothing
          Else
            'Erase the old one
        End If
        
        'Draw Each Flake
        If SnowFlakes(j).LargeFlake Then
            'First SRCAND the mask down
            BitBlt Picture1.hdc, SnowFlakes(j).X, SnowFlakes(j).Y, 29, 29, picSnowLarge(99).hdc, 29, 0, SRCAND
            'Then SRCCOPY the sprite
            BitBlt Picture1.hdc, SnowFlakes(j).X, SnowFlakes(j).Y, 29, 29, picSnowLarge(99).hdc, 0, 0, SRCPAINT
          Else
            'First SRCAND the mask down
            BitBlt Picture1.hdc, SnowFlakes(j).X, SnowFlakes(j).Y, 15, 15, picSnowSmall.hdc, 15, 0, SRCAND
            'Then SRCCOPY the sprite
            BitBlt Picture1.hdc, SnowFlakes(j).X, SnowFlakes(j).Y, 15, 15, picSnowSmall.hdc, 0, 0, SRCPAINT
        End If
    Next j

    Picture1.Refresh
    FPS = 1000 / (tmr1.Elapsed + 1)
    If BigFlakes Then
        NumSets = Max_Snow * 4
      Else
        NumSets = Max_Snow
    End If
    NumBlts = (Max_Flakes) * 2 + 1
    'Make sure we don't overflow the NumFrame var
    'At 30FPS that would take a little over
    '2 years. Heck, it doesn't hurt to code in for
    'Every possibility...
    If NumFrames > 2000000000 Then
        NumFrames = 0
        TempDBL = 0
    End If
    
    TempDBL = TempDBL + FPS
    NumFrames = NumFrames + 1
    AvgFPS = TempDBL / NumFrames
    
    If Not FullScreenSize Then
        Form1.Caption = "3D Snow - (" + Format(Int(FPS)) + ") FPS, " + Format(Int(AvgFPS)) + ") Average FPS, "
      Else
        Picture1.CurrentX = 615
        Picture1.CurrentY = 465
        Picture1.ForeColor = vbWhite
        Picture1.Print Format(Int(FPS))
    End If
    
    tmr1.StopTimer
    
    Exit Sub
E:
Unload Me
    End Sub



Sub InitSnow() 'Create Random Locations, Speed and Wind.
    
    For I = 1 To Max_Snow
        DoEvents
            Snow(I).X = Int(Rnd * Picture1.ScaleWidth)
            Snow(I).Z = Int(Rnd(1) * 5) + 1
            Snow(I).Y = Int(Rnd * Picture1.ScaleHeight)
            If RandSpeed Then
                Snow(I).Wind = Int(Rnd(1) * 3) + 1
                Snow(I).Speed = Int(Rnd(1) * 5) + 1
                Snow(I).CLR = GetSColor(Int(Rnd(1) * 5) + 1)
            Else
                Select Case Snow(I).Z
                    Case 1
                        Snow(I).Wind = 2
                        Snow(I).Speed = 2
                        Snow(I).CLR = GetSColor(1)
                    Case 2
                        Snow(I).Wind = 2
                        Snow(I).Speed = 3
                        Snow(I).CLR = GetSColor(2)
                    Case 3
                        Snow(I).Wind = 1
                        Snow(I).Speed = 2
                        Snow(I).CLR = GetSColor(3)
                    Case 4
                        Snow(I).Wind = 1
                        Snow(I).Speed = 2
                        Snow(I).CLR = GetSColor(4)
                    Case 5
                        Snow(I).Wind = 1
                        Snow(I).Speed = 1
                        Snow(I).CLR = GetSColor(5)
                End Select
            End If
        Next
    
    For I = 1 To Max_Flakes
        SnowFlakes(I).X = Int(Rnd * Picture1.ScaleWidth) - 50
        SnowFlakes(I).Y = Int(Rnd * Picture1.ScaleHeight)
        If Int(Rnd(1) * 2) + 1 = 1 Then
            SnowFlakes(I).LargeFlake = True
          Else
            SnowFlakes(I).LargeFlake = False
        End If
        If RandSpeed Then
            If SnowFlakes(I).LargeFlake Then
                SnowFlakes(I).Speed = Int(Rnd(1) * 7) + 1
                SnowFlakes(I).Wind = SnowFlakes(I).Speed - 1
              Else
                SnowFlakes(I).Speed = Int(Rnd(1) * 4) + 1
                SnowFlakes(I).Wind = SnowFlakes(I).Speed - 1
            End If
          Else
            If SnowFlakes(I).LargeFlake Then
                SnowFlakes(I).Speed = 3
                SnowFlakes(I).Wind = 2
              Else
                SnowFlakes(I).Speed = 2
                SnowFlakes(I).Wind = 1
            End If
        End If
        SnowFlakes(I).LastX = 99999
        SnowFlakes(I).lastY = 99999
        If SnowFlakes(I).LargeFlake Then
            SnowFlakes(I).FlakeNum = Int(Rnd(1) * 4)   '0 to 3
        End If
    Next I
    BitBlt Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture2.hdc, 0, 0, SRCCOPY
    End Sub

Sub cmdApply_Click()
    NumFrames = 0
    TempDBL = 0
    picSnowLarge(99).Picture = picSnowLarge(0).Picture
    picSnowLarge(99).Refresh
    If chkDots.Value = 1 Then BigFlakes = True Else BigFlakes = False
    If chkRandSpeed.Value = 1 Then RandSpeed = True Else RandSpeed = False
    If chkWrap.Value = 1 Then SnowWrap = True Else SnowWrap = False
    If chkRotate.Value = 1 Then RotatingFlakes = True Else RotatingFlakes = False
    'If chkSet.Value = 1 Then setPixelIt = True Else setPixelIt = False
    setPixelIt = True
    
    If Val(txtDots.Text) > 9999 Or Val(txtDots.Text) < 0 Then txtDots.Text = "500"
    If Val(txtFlakes.Text) > 9999 Or Val(txtFlakes.Text) < 0 Then txtFlakes.Text = "50"
    If Val(txtAngle.Text) > 360 Or Val(txtAngle.Text) < -360 Then txtAngle.Text = "3"
    
    FLAKEROTATEANGLE = Val(txtAngle.Text)
    Max_Snow = Val(txtDots.Text)
    Max_Flakes = Val(txtFlakes.Text)
    
    If Not ConfigMode Then
        Call cmdFullScreen_Click
    End If
    
    InitSnow
    Looper
End Sub

Sub cmdFullScreen_Click()
    'If it's fullscreen, make it normal size, and vise - versa.
    If FullScreenSize Then
        Call DisableCtrlAltDelete(False)
        FullScreenSize = False
        Call SetScreen(False)
        Form1.Caption = "3D Snow"
        Form1.Height = NORMSCREENHEIGHT
        cmdFullScreen.Caption = "&Full Screen"
        Form1.Top = 0
        Form1.Left = 0
        Form1.SetFocus
      Else
        Call DisableCtrlAltDelete(True)
        FullScreenSize = True
        Call SetScreen(True)
        Form1.Caption = ""
        Form1.Height = FULLSCREENHEIGHT
        cmdFullScreen.Caption = "&Original Size"
        Form1.Top = 0
        Form1.Left = 0
        Form1.SetFocus
    End If
End Sub

Private Sub cmdQuit_Click()
    'Self explanatory...
    Ended = True
    DoEvents
    Call DisableCtrlAltDelete(False)
    Unload Me
    End
    
End Sub

Private Sub cmdQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'This is just to make VB think a little, to help solve a bug of it's - not mine
    'Where you can't click close right after you switch back from full screen
    'Sort of odd, but it's not my fault. This seems to work however.
    DoEvents
End Sub

Private Sub Form_Load()
   
    'Load the available background options
    GetBackgrounds
    'Find out what the user is running their monitor at right now
    GetScreenInfo
    'Initialize random numbers.
    Randomize Timer
    'Show the main form.
    Me.Show
    'Apply the settings (Sorta cheating, but it makes it work better with less code this way).
    Call cmdApply_Click
    'Make sure the 2 pictures match (Or else it would not look very good)
    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
   
    End Sub

Sub Looper()
    'Main loop for the program.
    'Loops until the cows come home, or ENDED=TRUE, Whichever happens first.
    DoEvents
    Do
        DoEvents
        DoSnow
    Loop While Not Ended

End Sub

Function GetSColor(Z As Integer)
    'For the large flakes, this makes their color emulate
    'a 3D look.
    'The Z axis is emulated by color.
    '(distant snow is darker)
    If BigFlakes Then
        Select Case Z
            Case 1
            GetSColor = vbWhite
            Case 2
            GetSColor = vbWhite
            Case 3
            GetSColor = &HE0E0E0
            Case 4
            GetSColor = &HC0C0C0
            Case 5
            GetSColor = &HC0C0C0
            Case Else
            GetSColor = vbWhite
        End Select
    Else
        GetSColor = vbWhite
    End If

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SSMode Then End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Tell Windows that it's not in a screensaver mode
    Call DisableCtrlAltDelete(False)
    DoEvents
    'Makes sure we don't stick the user in 640x480.
    If FullScreenSize Then
        Call SetScreen(False)
    End If
    Ended = True
    DoEvents
    Unload Me
    End
    
End Sub

Private Sub HScroll1_Change()
    Picture2.Picture = LoadPicture(App.Path & "\" + Backgrounds(HScroll1.Value))
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Just makes the program a little user-friendlier.
    If X > Picture1.ScaleWidth Then Exit Sub
    If Y > Picture1.ScaleHeight Then Exit Sub
    'I really shouldn't have to check those, but ATI's video drivers are BUGGY!!!
    'Why should a picturebox_Mousedown event be called after a screen resolution change
    'Even when the user didn't click on the picturebox, but ANYWHERE ELSE??????????????
    'Then again, it only happens on one of my systems, and nowhere else. Damn ATI Video cards...
    DoEvents
    Call cmdFullScreen_Click

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static mouse As Integer
    If Not SSMode Then Exit Sub
    mouse = mouse + 1
    
    If mouse > 8 Then Call cmdQuit_Click
End Sub

Private Sub tmrFlakeHorzRotate_Timer()
    Static ANG As Integer
    If Max_Flakes < 1 Then Exit Sub
    If Not RotatingFlakes Then Exit Sub
    ANG = ANG + FLAKEROTATEANGLE
    If ANG >= 360 Then ANG = 0
    Call RotateFlakes(ANG)
    'Basically, this just rotates the large snowflake 3 degrees counter-clockwise
    'For every call of the timer.
End Sub


Sub SetScreen(Full As Boolean)
    'Thanks go to www.vb-world.net for this tip also.
    'Sets the screen to 640x480 if full=TRUE, otherwise back to the original res if FALSE.
    
    Dim DevM As DEVMODE
    Dim erg As Long
    Dim an As Variant
    erg& = EnumDisplaySettings(0&, 0&, DevM)
    If Full Then
        ShowCursor (False)
        DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
        DevM.dmPelsWidth = 640 'ScreenWidth
        DevM.dmPelsHeight = 480 'ScreenHeight
        
        erg& = ChangeDisplaySettings(DevM, CDS_TEST)
        
        Select Case erg&
        Case DISP_CHANGE_SUCCESSFUL
            erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
        Case Else
            MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
        End Select
      Else
        ShowCursor (True)
        DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
        DevM.dmPelsWidth = OrgScreenX 'ScreenWidth
        DevM.dmPelsHeight = OrgScreenY 'ScreenHeight
        erg& = ChangeDisplaySettings(DevM, CDS_TEST)
        Select Case erg&
        Case DISP_CHANGE_SUCCESSFUL
            erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
        Case Else
            MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
        End Select
    End If
End Sub

Sub GetScreenInfo()
    'Remembers what the res was when the program started, as not to PO the user...
    OrgScreenX = GetSystemMetrics(SM_CXSCREEN)
    OrgScreenY = GetSystemMetrics(SM_CYSCREEN)
End Sub

Public Sub RotatePicture(ByVal SourcehDC As Long, ByVal DesthDC As Long, ByVal AngleInRadians As Double, ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, ByVal OrigX As Integer, ByVal OrigY As Integer, ByVal NewX As Integer, ByVal NewY As Integer)
    'Thanks go to www.vb-world.net for this tip.

    Dim sin_theta As Double
    Dim cos_theta As Double
    Dim MinX As Integer
    Dim MaxX As Integer
    Dim MinY As Integer
    Dim MaxY As Integer
    Dim tx As Integer
    Dim ty As Integer
    Dim fx As Double
    Dim fy As Double
    Dim ifx As Integer
    Dim ify As Integer
    
    ' Compute the sine and cosine of theta.
    sin_theta = Sin(AngleInRadians)
    cos_theta = Cos(AngleInRadians)
    
    ' Make some bounds for new picture
    MinX = (Left - OrigX) * cos_theta + (Top - OrigY) * sin_theta + NewX
    MinY = -(Left - OrigX) * sin_theta + (Top - OrigY) * cos_theta + NewY
    MaxX = MinX
    MaxY = MinY
    
    tx = (Left - OrigX) * cos_theta + (Bottom - OrigY) * sin_theta + NewX
    ty = -(Left - OrigX) * sin_theta + (Bottom - OrigY) * cos_theta + NewY
    If MinX > tx Then MinX = tx
    If MinY > ty Then MinY = ty
    If MaxX < tx Then MaxX = tx
    If MaxY < ty Then MaxY = ty
    
    tx = (Right - OrigX) * cos_theta + (Top - OrigY) * sin_theta + NewX
    ty = -(Right - OrigX) * sin_theta + (Top - OrigY) * cos_theta + NewY
    If MinX > tx Then MinX = tx
    If MinY > ty Then MinY = ty
    If MaxX < tx Then MaxX = tx
    If MaxY < ty Then MaxY = ty
    
    tx = (Right - OrigX) * cos_theta + (Bottom - OrigY) * sin_theta + NewX
    ty = -(Right - OrigX) * sin_theta + (Bottom - OrigY) * cos_theta + NewY
    If MinX > tx Then MinX = tx
    If MinY > ty Then MinY = ty
    If MaxX < tx Then MaxX = tx
    If MaxY < ty Then MaxY = ty
    
    If MinX < 1 Then MinX = 1
    If MaxX < 1 Then MaxX = 1
    
    If MinY < 1 Then MinY = 1
    If MaxY < 1 Then MaxY = 1
    
    ' Perform the rotation.
    For ty = MinY To MaxY
    For tx = MinX To MaxX
    
    ' Find the location (fx, fy) that maps to the pixel (tx, ty).
    fx = (tx - NewX) * cos_theta - (ty - NewY) * sin_theta + OrigX
    fy = (tx - NewX) * sin_theta + (ty - NewY) * cos_theta + OrigY
    
    ify = Fix(fy)
    ifx = Fix(fx)
    If ifx >= Left And ifx < Right And ify >= Top And ify < Bottom Then
    Call SetPixelV(DesthDC, tx, ty, GetPixel(SourcehDC, ifx, ify))
    End If
    Next tx
    Next ty

End Sub


Sub RotateFlakes(Angle As Integer)
    'Just calles the picture rotation procedure
    
    Dim theta As Double
    Dim CurrFlake As Integer
    
    Dim FlakeRotNum As Byte
    Dim FromHDC As Long
    Dim ToHDC As Long
    Dim FromNum  As Integer
    Dim ToNum As Integer
    
    lblAngle.Caption = Angle
    FromNum = 0
    ToNum = 99
    theta = Pi * (Angle) / 180
    picSnowLarge(ToNum).Cls
    
    RotatePicture picSnowLarge(FromNum).hdc, picSnowLarge(ToNum).hdc, _
    theta, 0, 0, ((picSnowLarge(FromNum).ScaleWidth) / 2) - 1, _
    picSnowLarge(FromNum).ScaleHeight - 1, ((picSnowLarge(FromNum).ScaleWidth) / 2) / 2, _
    picSnowLarge(FromNum).ScaleHeight / 2, ((picSnowLarge(ToNum).ScaleWidth) / 2) / 2, _
    picSnowLarge(ToNum).ScaleHeight / 2
    
    
    MakeMask (ToNum)
    picSnowLarge(ToNum).Refresh
End Sub

Sub MakeMask(I As Integer)
    'This is my own on the fly mask generator, since none I found could do what I wanted.
    'This procedure is from my castle game - www.comp-info.net/castle
    'Slightly modified to make it work with this program.
    
    Dim X As Long
    Dim Y As Long
    Dim MaskColor As Long
    
    MaskColor = vbBlack
    
    For X = 0 To picSnowLarge(I).ScaleWidth - 1 Step 1
        For Y = 0 To picSnowLarge(I).ScaleHeight - 1 Step 1
            If GetPixel(picSnowLarge(I).hdc, X, Y) = MaskColor Then
                SetPixel picSnowLarge(I).hdc, X + Int(picSnowLarge(I).ScaleWidth / 2), Y, vbWhite
              Else
                SetPixel picSnowLarge(I).hdc, X + Int(picSnowLarge(I).ScaleWidth / 2), Y, vbBlack
            End If
        Next
    Next
    picSnowLarge(I).Refresh
End Sub

Sub GetBackgrounds()
    Dim FF As Byte
    Dim I As Integer
    Dim NumBackgrounds As Integer
    
    FF = FreeFile
    Close   'just to be sure
    Open App.Path & "\backgrnd.dat" For Input As #FF
        Input #FF, NumBackgrounds
        For I = 1 To NumBackgrounds
            Input #FF, Backgrounds(I)
        Next I
    Close #FF
    Close   'Never hurts
    
    HScroll1.Max = NumBackgrounds
    
End Sub

Sub DisableCtrlAltDelete(bDisabled As Boolean)
    'Thanks go to www.vb-world.net for this one also.
    Dim X As Long
    X = SystemParametersInfo(97, bDisabled, CStr(1), 0)
End Sub

