VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dr. Mario"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7185
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmain.frx":0442
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picoptions 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5160
      Left            =   0
      Picture         =   "frmmain.frx":36C6
      ScaleHeight     =   344
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   0
      Width           =   7200
      Begin VB.Image picmenu 
         Height          =   5160
         Left            =   0
         Picture         =   "frmmain.frx":51A1
         Top             =   0
         Width           =   7200
      End
      Begin VB.Image imgmenu 
         Height          =   600
         Index           =   2
         Left            =   960
         Picture         =   "frmmain.frx":8D26
         Top             =   3720
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.Image imgmenu 
         Height          =   600
         Index           =   1
         Left            =   960
         Picture         =   "frmmain.frx":8FC0
         Top             =   2100
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Image imgmenu 
         Height          =   600
         Index           =   0
         Left            =   960
         Picture         =   "frmmain.frx":91C2
         Top             =   630
         Width           =   3090
      End
      Begin VB.Image imgmusicselection 
         Height          =   615
         Left            =   1440
         Top             =   4320
         Width           =   4455
      End
      Begin VB.Image Imgmusic 
         Height          =   480
         Left            =   1500
         Picture         =   "frmmain.frx":9481
         Top             =   4410
         Width           =   1500
      End
      Begin VB.Image imgspeed 
         Height          =   615
         Left            =   2400
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Image imgviruslevel 
         Height          =   660
         Left            =   2430
         Top             =   1320
         Width           =   2490
      End
      Begin VB.Image imgbig 
         Height          =   180
         Left            =   3600
         Picture         =   "frmmain.frx":9530
         Top             =   2730
         Width           =   630
      End
      Begin VB.Image imgsmall 
         Height          =   210
         Left            =   2355
         Picture         =   "frmmain.frx":958C
         Top             =   1350
         Width           =   210
      End
      Begin VB.Label lblmain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   5280
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.PictureBox Picfuture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3375
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   3840
      Left            =   2640
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   0
      Top             =   1080
      Width           =   1920
      Begin VB.Shape shpstart 
         BackColor       =   &H80000008&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   240
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Image imglevelclear 
         Height          =   3840
         Left            =   0
         Picture         =   "frmmain.frx":95DE
         Top             =   0
         Visible         =   0   'False
         Width           =   1920
      End
   End
   Begin VB.PictureBox picskin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      Picture         =   "frmmain.frx":A097
      ScaleHeight     =   960
      ScaleWidth      =   1920
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   1920
      Begin VB.Timer Timersmall 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   360
         Top             =   600
      End
      Begin VB.Timer Timermain 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   0
         Top             =   600
      End
   End
   Begin VB.Label lblpaused 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "PAUSED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Shape shpmain 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      Height          =   135
      Left            =   6240
      Top             =   1080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblmain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   5280
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblmain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5280
      TabIndex        =   5
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label lblmain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   4
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblmain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblmain 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000010000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Initial As Long = 50, Factor As Double = 0.9
Public WithEvents DrMario As DrMarioCLS, Virus As Long, cX As Single, cY As Single, isdown As Boolean
Attribute DrMario.VB_VarHelpID = -1
Dim currmenu As Long, VirusLevel As Long, Speed As Long, Music As Long, Score As Long

Private Sub DrMario_LostGame()
    Virus = 4
    If MsgBox("Would you like to try again?", vbYesNo, "Game Over") = vbYes Then
        picmenu.Visible = True
        picoptions.Visible = True
        Timermain.Enabled = False
        Timersmall.Enabled = False
        Form_Load
    Else
        End
    End If
End Sub

Private Sub DrMario_PulledTile()
    Static Count As Long
    Count = Count + 1
    Score = DrMario.Score(Speed, 1)
    If Count = 10 Then
        If Timersmall.Interval * Factor > 0 Then Timersmall.Interval = Timersmall.Interval * Factor
        Count = 0
    End If
End Sub

Private Sub DrMario_VirusChanged(Virii As Long)
    lblmain(4) = Virii
    
    lblmain(1) = Format(Val(lblmain(1)) + Score, "0000000000")
    If Score > Val(lblmain(0)) Then lblmain(0) = lblmain(1)
End Sub

Private Sub DrMario_VirusKilled(number As Long)
    Score = DrMario.Score(Speed, number)
End Sub

Private Sub DrMario_WonGame()
    picmain.Cls
    Picfuture.Visible = False
    Timersmall.Enabled = False
    imglevelclear.Visible = True
    PlayMusic "VICTORY"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 80, 19, 145 'P
            shpmain.Visible = Timermain.Enabled
            lblpaused.Visible = Timermain.Enabled
            
            Timermain.Enabled = Not Timermain.Enabled
            Timersmall.Enabled = Timermain.Enabled
            If Music < 2 Then MediaPlay
            picmain.Visible = Timermain.Enabled
            Picfuture.Visible = Timermain.Enabled
        Case Else
            If imglevelclear.Visible Then
                imglevelclear_Click
            Else
                DrMario.KeyDown KeyCode
            End If
    End Select
End Sub

Private Sub Form_Load()
    picoptions.Move 0, 0
    picmenu.Move 0, 0
    Speed = 1
    Set DrMario = New DrMarioCLS
    With DrMario
        .InitDrMario picskin, picmain, Picfuture
        .RandomizeBoard 4, 7
        Virus = 4
    End With
    shpmain.Move 0, 0, Width / 15, Height / 15
    lblmain(0) = GetSetting("Dr. Mario", "Main", "High Score", "0000010000")
    ALIAS = "DRMARIO"
    PlayMusic "TITLE"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MediaCloseAll
    SaveSetting "Dr. Mario", "Main", "High Score", lblmain(0)
End Sub

Private Sub imglevelclear_Click()
    lblmain(2) = Virus - 2
    Virus = Virus + 1
    Timersmall.Interval = getInterval
    Timersmall.Enabled = True
    Picfuture.Visible = True
    shpstart.Visible = False
    imglevelclear.Visible = False
    DoMusic
End Sub

Private Sub imgmusicselection_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    isdown = True
    currmenu = 2
    ResetMenu
    imgmusicselection_MouseMove Button, Shift, x, y
End Sub

Private Sub imgmusicselection_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If isdown Then
        If x < 1635 Then
            Imgmusic.Left = 100
            Music = 0
        ElseIf x < 3285 Then
            Imgmusic.Left = 212
            Music = 1
        Else
            Imgmusic.Left = 308
            Music = 2
        End If
    End If
End Sub

Private Sub imgmusicselection_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isdown = False
End Sub

Private Sub imgspeed_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    isdown = True
    currmenu = 1
    ResetMenu
    imgspeed_MouseMove Button, Shift, x, y
End Sub

Private Sub imgspeed_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If isdown Then
    Select Case x
        Case Is <= 705
            lblmain(3) = "LOW"
            imgbig.Left = 162
            Speed = 0
        Case Is >= 2370
            lblmain(3) = "HI"
            imgbig.Left = 315
            Speed = 2
        Case Else
            If x >= 1185 And x <= 1905 Then
                lblmain(3) = "MED"
                imgbig.Left = 240
                Speed = 1
            End If
    End Select
    End If
End Sub

Private Sub imgspeed_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isdown = False
End Sub

Private Sub imgviruslevel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    isdown = True
    currmenu = 0
    ResetMenu
    imgviruslevel_MouseMove Button, Shift, x, y
End Sub

Private Sub imgviruslevel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If isdown Then
        If x < 0 Then x = 0
        x = x / 15
        If x > imgviruslevel.Width Then x = imgviruslevel.Width
        VirusLevel = x / (imgviruslevel.Width / 20)
        lblmain(5) = Format(VirusLevel, "00")
        imgsmall.Left = x + imgviruslevel.Left - imgsmall.Width / 2
    End If
End Sub

Private Sub imgviruslevel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isdown = False
End Sub

Private Sub picmain_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub picmain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    isdown = True
    cX = x
    cY = y
End Sub

Public Sub PlayMusic(File As String, Optional extention As String = ".mid")
    MediaLoad Replace(App.Path & "\" & File & extention, "\\", "\")
    MediaPlay
End Sub


Private Sub picmain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Const modulous As Long = 16
    If isdown Then
        Select Case cX \ modulous
            Case Is < x \ modulous: DrMario.DoAction moveright
            Case Is > x \ modulous: DrMario.DoAction MoveLeft
        End Select
        Select Case cY \ modulous
            Case Is < y \ modulous: DrMario.DoAction movedown
        End Select
        cX = x
        cY = y
    End If
End Sub

Private Sub picmain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
isdown = False
If Button = vbRightButton Then DrMario.DoAction RotateTiles
End Sub

Private Sub picmenu_Click()
    picmenu.Visible = False
    PlayMusic "OPTIONS"
    picoptions.SetFocus
End Sub

Private Sub picmenu_KeyPress(KeyAscii As Integer)
    picmenu_Click
End Sub

Private Sub picoptions_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, 32
            picoptions.Visible = False
            Me.SetFocus
            DoMusic
            DrMario.RandomizeBoard (VirusLevel + 2) * 2, 12
            lblmain(2) = VirusLevel
            Timersmall.Interval = getInterval
            Timersmall.Enabled = True
            Timermain.Enabled = True
        Case 38 'up
            currmenu = currmenu - 1
            If currmenu < 0 Then currmenu = 2
            ResetMenu
        Case 40 'down
            currmenu = currmenu + 1
            If currmenu > 2 Then currmenu = 0
            ResetMenu
        Case 37 'left
            Select Case currmenu
                Case 0: If VirusLevel > 0 Then SetVirusLevel VirusLevel - 1
                Case 1: SetVirusSpeed Speed - 1
                Case 2: SetMusic Music - 1
            End Select
        Case 39 'right
            Select Case currmenu
                Case 0: If VirusLevel < 20 Then SetVirusLevel VirusLevel + 1
                Case 1: SetVirusSpeed Speed + 1
                Case 2: SetMusic Music + 1
            End Select
    End Select
End Sub

Private Sub DoMusic()
            MediaStop
            If Music = 0 Then PlayMusic "FEVER"
            If Music = 1 Then PlayMusic "CHILL"
End Sub

Private Sub SetVirusLevel(number As Long)
    isdown = True
    imgviruslevel_MouseMove 0, 0, (number * (imgviruslevel.Width / 20)) * 15, 0
    isdown = False
End Sub

Private Sub SetVirusSpeed(ByVal number As Long)
    Do Until number >= 0
        number = number + 3
    Loop
    Do Until number <= 2
        number = number - 3
    Loop
    isdown = True
    Select Case number
        Case 0: imgspeed_MouseMove 0, 0, 705, 0
        Case 1: imgspeed_MouseMove 0, 0, 1185, 0
        Case 2: imgspeed_MouseMove 0, 0, 2370, 0
    End Select
    isdown = False
End Sub
Private Sub SetMusic(ByVal number As Long)
    Do Until number >= 0
        number = number + 3
    Loop
    Do Until number <= 2
        number = number - 3
    Loop
    isdown = True
    Select Case number
        Case 0: imgmusicselection_MouseMove 0, 0, 1630, 0
        Case 1: imgmusicselection_MouseMove 0, 0, 3280, 0
        Case 2: imgmusicselection_MouseMove 0, 0, 3290, 0
    End Select
    isdown = False
End Sub

Private Sub Timermain_Timer()
    If imglevelclear.Visible Then
        shpstart.Visible = Not shpstart.Visible
    Else
        DrMario.SwitchViral
    End If
    If MediaTimeRemaining = 0 Then MediaSeekto
End Sub

Private Sub Timersmall_Timer()
    If DrMario.TileCount = 0 Then
        DrMario.RandomizeBoard Virus, 7
    Else
        DrMario.DecrementOffsets 4
        DrMario.DecrementCurrentTile
        DrMario.DrawScreen
    End If
End Sub

Public Function getInterval() As Long
    Select Case Speed
        Case 0: getInterval = 100
        Case 1: getInterval = 50
        Case 2: getInterval = 25
    End Select
End Function

Public Sub ResetMenu()
    Dim temp As Long
    For temp = 0 To 2
        imgmenu(temp).Visible = temp = currmenu
    Next
End Sub
