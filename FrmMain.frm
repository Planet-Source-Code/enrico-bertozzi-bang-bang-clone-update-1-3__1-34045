VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Bang Bang Clone!"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Back 
      AutoRedraw      =   -1  'True
      Height          =   5055
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   240
      Width           =   6975
      Begin VB.PictureBox pExplsrc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pExplmsk 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree2msk 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   1800
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree2src 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree1msk 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   840
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox pTree1src 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   360
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picCopy 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3840
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox picMask 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   3360
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.PictureBox picSrc 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   300
         Left            =   2880
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start game"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   4560
         Width           =   1335
      End
      Begin VB.CommandButton cmdRegenerate 
         Caption         =   "Regenerate"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   4560
         Width           =   1335
      End
      Begin VB.PictureBox pWind 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2640
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   6
         Top             =   4320
         Visible         =   0   'False
         Width           =   1695
         Begin VB.Shape Shape1 
            BorderColor     =   &H0000FF00&
            Height          =   375
            Left            =   0
            Top             =   0
            Width           =   1695
         End
         Begin VB.Line prWind 
            BorderColor     =   &H000000FF&
            Index           =   1
            Visible         =   0   'False
            X1              =   48
            X2              =   48
            Y1              =   0
            Y2              =   32
         End
         Begin VB.Line prWind 
            BorderColor     =   &H00FFC0C0&
            Index           =   2
            Visible         =   0   'False
            X1              =   80
            X2              =   80
            Y1              =   0
            Y2              =   32
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            X1              =   57
            X2              =   57
            Y1              =   0
            Y2              =   24
         End
      End
      Begin VB.Line TgtLX 
         Visible         =   0   'False
         X1              =   294
         X2              =   322
         Y1              =   208
         Y2              =   208
      End
      Begin VB.Line TgtLY 
         Visible         =   0   'False
         X1              =   308
         X2              =   308
         Y1              =   194
         Y2              =   222
      End
      Begin VB.Shape AimTgt 
         BorderWidth     =   2
         Height          =   315
         Left            =   4440
         Shape           =   3  'Circle
         Top             =   2970
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Line AimLine 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   208
         X2              =   272
         Y1              =   208
         Y2              =   208
      End
      Begin VB.Label lWind 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Wind"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lAct 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press Space"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Image imgP2 
         Height          =   135
         Left            =   1680
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgP1 
         Height          =   135
         Left            =   1320
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lsInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmMain.frx":0442
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.Label lP2Score 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   6975
   End
   Begin VB.Label lP1Score 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Const Zoom% = 100 'PLEASE LEAVE 100 IN THERE OR GAME MAY FREEZE!

Private Sub cmdRegenerate_Click()

Dim I%, PndStart%, CurrTerLv%, PrevTerLv%, TerStartLvP%, PndFreq%, PndStop%, PndNormalize!, Invert&, PndAmplif!, PndLen!, BaseXRange%, CTreePos%

On Error GoTo Restart
Restart:

StClear 'clear the vars
Back.Cls 'clear picturebox
VarResetFade 'create random gradient in picturebox

Invert = Back.Point(1, 1) 'we'll use this color to get the inverted color for text to contrast better
CurrTerLv = Int(Rnd * 7 / 8 * Back.ScaleHeight) + (1 / 8 * Back.ScaleHeight) '<< random landscape height start level
PrevTerLv = CurrTerLv 'initial terrain level before change
TerStartLvP = CurrTerLv 'last terrain level
PndFreq = Int(Rnd * 1 / 12 * Back.ScaleHeight) '<< 1/12 (fraction) = complexity, lower to increase, modify also the below one...
BaseXRange = Back.ScaleWidth / 2 - 60
P1BaseEnd = Rnd * BaseXRange + 60 'calculate cannon positions
P1BaseStart = P1BaseEnd - 40
P2BaseEnd = Rnd * BaseXRange + Back.ScaleWidth / 2 + 40
P2BaseStart = P2BaseEnd - 40

For I = 1 To Back.ScaleWidth
'I use a COS function to generate a random landscape
'with randomized parameters (amplitude, lenght)
    If I = PndFreq Then

        PndStart = I
Exceeds:
        PndAmplif = Rnd * 2 - 1 '<< amplitude
        PndLen = Rnd * 2 + 0.5 '<< lenght
        PndNormalize = Zoom * PndAmplif 'we will subtract this to normalize to the start level
        If (TerLvl(I - 1) - 2 * PndNormalize > Me.ScaleHeight) Or (TerLvl(I - 1) - 2 * PndNormalize < 3 / 16 * Me.ScaleHeight) Then GoTo Exceeds '<< lower and higher mountain bound
        PndStop = I + Int(180 / PndLen)

        Do 'CYCLE START
            If (I > P1BaseStart And I < P1BaseEnd) Or (I > P2BaseStart And I < P2BaseEnd) Then
                TerStartLvP = CurrTerLv 'if we are drawing a cannon base, the height is constant
            Else
                TerStartLvP = CurrTerLv 'if not, we increment the initial level with the function
                CurrTerLv = Int(PrevTerLv + Zoom * (PndAmplif * Cos(PndLen * ((I - PndStart) * 3.141592 / 180)))) - PndNormalize
            End If

            Back.Line (I, CurrTerLv + 1)-(I, Back.ScaleWidth), RGB(0, 170, 0) '<< terrain color, modify also the below one
            Back.Line (I - 1, TerStartLvP)-(I, CurrTerLv), RGB(0, 0, 0)
            TerLvl(I) = CurrTerLv 'set the main variable with that level
            I = I + 1
        Loop Until I = PndStop 'Loop until we finish drawing the terrain level change

        Back.Line (I, CurrTerLv + 1)-(I, Back.ScaleWidth), RGB(0, 170, 0) '<< terrain color, modify also the below one
        Back.Line (I - 1, CurrTerLv)-(I + 1, CurrTerLv), RGB(0, 0, 0)
        PndFreq = Int(Rnd * 1 / 12 * Back.ScaleHeight) + I '<< 1/12 (fraction) = complexity, lower to increase
        PrevTerLv = CurrTerLv
        TerLvl(I) = CurrTerLv

    Else

        Back.PSet (I, CurrTerLv), RGB(0, 0, 0)
        Back.Line (I, CurrTerLv + 1)-(I, Back.ScaleWidth), RGB(0, 170, 0) '<< terrain color
        TerLvl(I) = CurrTerLv
    End If

Next I

NumTrees = Int(Rnd * 10) + 5 'how many trees

For I = 1 To NumTrees
    CTreePos = Int(Rnd * (Back.ScaleWidth - 48)) + 24 'randomize tree pos
    'draw one of the two trees
    If Int(Rnd * 2) = 0 Then DrwTranspSpriteBlt Back, CTreePos - 24, TerLvl(CTreePos) - 46, pTree1src, pTree1msk Else DrwTranspSpriteBlt Back, CTreePos - 24, TerLvl(CTreePos) - 46, pTree2src, pTree2msk
Next I

'ALL these instructions exist only to position objects in form,
'to hide/unhide them and to set its colors
InvertR = LongToR(Back.Point(2, 2))
InvertG = LongToG(Back.Point(2, 2))
InvertB = LongToB(Back.Point(2, 2))
lsInfo.ForeColor = RGB(256 - InvertR, 256 - InvertG, 256 - InvertB)
AllZero = False 'We have regenerated the landscape, so allow starting the game
StartX(1) = P1BaseEnd - 7
StartX(2) = P2BaseStart + 7
StartY(1) = TerLvl(P1BaseEnd - 1) - 32
StartY(2) = TerLvl(P2BaseStart + 1) - 32
imgP1.Top = TerLvl(P1BaseStart + 1) - 32: imgP1.Left = P1BaseStart + 1
imgP2.Top = TerLvl(P2BaseEnd - 1) - 32: imgP2.Left = P2BaseStart + 7
imgP1.Visible = True
imgP2.Visible = True
lAct.ForeColor = RGB(256 - InvertR, 256 - InvertG, 256 - InvertB)
temp& = Point(P1BaseStart, TerLvl(P1BaseStart) - 10)
'Not sure if this'll be correct
'AimerColor = IIf((LongToR(temp) + LongToG(temp) + LongToB(temp)) / 3 > 127, 0, RGB(255, 255, 255))
AimerColor = RGB(255, 255, 255) 'use this for now

On Error GoTo 0

End Sub

Private Sub cmdStart_Click()
'IF user hasn't pressed the Regenerate cmdbutton, remark it.
If AllZero Then MsgBox "Click 'Regenerate' first!", vbExclamation: Exit Sub

If DispKeys Then
    MsgBox "Game keys:" & vbCr & vbCr & "up/down arrow" & vbTab & "aim up/down" & vbCr & "keypad +/-" & vbTab & "increase/decrease power" & vbCr & vbCr & "press ESC to stop game without exiting" _
    & vbCr & vbCr & "In the windbox, you can see lines, indicating your previous wind value and helping you to shoot properly", vbInformation, "Keys"
    DispKeys = False
End If

CanShoot = True
PlrShoot(1).Angle = 45
PlrShoot(1).Power = 100
PlrShoot(2).Angle = 45
PlrShoot(2).Power = 100
PlrShoots(1) = 0
PlrShoots(2) = 0

If Not InGame Then 'If we are starting a NEW game, randomize the turn
    TurnOf = Int(Rnd * 2) + 1
Else 'If not, make start first who previously lose the match
    If HasWon = 1 Then TurnOf = 2 Else TurnOf = 1
End If

HasWon = 0
InGame = True
WaitSpace = True
ObjSet osPlaying
Wind = 0.07 * Rnd - 0.035 'set starting wind
PrevWind(1) = Wind
PrevWind(2) = Wind

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape And InGame Then
    If StopGame Then

        InGame = False
        FrmMain.Back.Cls 'clear
        StClear
        VarResetFade 'create a new gradient
        ObjSet osWelcome 'unhide all the Welcome control set
        AllZero = True 'users has to click Regenerate
        CanShoot = False 'user can't shoot
        PlrScore(1) = 0 'reset score
        PlrScore(2) = 0
        cmdRegenerate_Click 'simulate a click in the Regenerate cmdbutton
    
    End If
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

Dim WndIncr!
Dim AnglBuf%, PwrBuf%
Dim KeyBufClr&

'if I pressed space and I can shoot and I am ingame, and I'm not in aiming phase
If KeyAscii = 32 And CanShoot And InGame And WaitSpace Then
'>start ---------------------- Draw Windbox
    WaitSpace = False
    StClear
    lAct.Visible = False
ReRnd:
    Randomize (Timer)
    WndIncr = 0.004 * Rnd - 0.002
    'increment wind
    If Wind + WndIncr < -0.035 Or Wind + WndIncr > 0.035 Then GoTo ReRnd
    Wind = Wind + WndIncr

    If PlrShoots(1) > 0 And PlrShoots(2) > 0 Then
        prWind(TurnOf).Visible = True
        prWind(IIf(TurnOf = 1, 2, 1)).Visible = False
    End If

    prWind(TurnOf).X1 = Int((PrevWind(TurnOf) + 0.035) * pWind.Width / 0.07)
    prWind(TurnOf).X2 = prWind(TurnOf).X1
    PrevWind(TurnOf) = Wind
    pWind.Cls
    R(1) = 0 'set color variables to create a predefined color
    G(1) = 196 'gradient pattern.
    B(1) = 0 'then we'll paint white unused areas
    R(2) = 0
    G(2) = 128
    B(2) = 0
    R(3) = 0
    G(3) = 128
    B(3) = 0
    R(4) = 0
    G(4) = 196
    B(4) = 0
    FStep(1) = 0
    FStep(2) = pWind.Width / 2
    FStep(3) = pWind.Width / 2
    FStep(4) = pWind.Width
    ObjFade pWind, drHorizontal
    Select Case Wind
    Case Is > 0
        For I = 0 To pWind.Width / 2
            pWind.Line (I, 0)-(I, pWind.Height), 16384
        Next I

        For I = Int((PrevWind(TurnOf) + 0.035) * pWind.Width / 0.07) To pWind.Width
            pWind.Line (I, 0)-(I, pWind.Height), 16384
        Next I

    Case Is < 0
        For I = 0 To Int((PrevWind(TurnOf) + 0.035) * pWind.Width / 0.07)
            pWind.Line (I, 0)-(I, pWind.Height), 16384
        Next I

        For I = Int(pWind.Width / 2) To pWind.Width
            pWind.Line (I, 0)-(I, pWind.Height), 16384
        Next I
    Case Else
    End Select
'>end -------------------------- Draw Windbox
'>start ------------------------ AIMING
    If TurnOf = 1 Then
        AimStartX = P1BaseEnd - 15 'set starting AIMER position
        AimStartY = TerLvl(P1BaseEnd - 1) - 27
    Else
        AimStartX = P2BaseStart + 15
        AimStartY = TerLvl(P2BaseStart + 1) - 27
    End If

    TVisible = True 'make it visible

    AnglBuf = 0 'these are used not to redraw all the time the AIMER that can
    PwrBuf = 0 'then appear flashy
    
    'empty the keybuffer, otherwise first shoot will start immediately (!)
    KeyBufClr = GetAsyncKeyState(vbKeyReturn)
    KeyBufClr = GetAsyncKeyState(vbKeyReturn)

    Do
        'increment/decrement angle and power
        If PlrShoot(TurnOf).Angle < 90 And GetAsyncKeyState(vbKeyUp) Then PlrShoot(TurnOf).Angle = PlrShoot(TurnOf).Angle + 1
        If PlrShoot(TurnOf).Angle > 0 And GetAsyncKeyState(vbKeyDown) Then PlrShoot(TurnOf).Angle = PlrShoot(TurnOf).Angle - 1
        If PlrShoot(TurnOf).Power < 150 And GetAsyncKeyState(vbKeyAdd) Then PlrShoot(TurnOf).Power = PlrShoot(TurnOf).Power + 1
        If PlrShoot(TurnOf).Power > 0 And GetAsyncKeyState(vbKeySubtract) Then PlrShoot(TurnOf).Power = PlrShoot(TurnOf).Power - 1
        
        'show the AIMER in correct position
        If AnglBuf <> PlrShoot(TurnOf).Angle Or PwrBuf <> PlrShoot(TurnOf).Power Then TgtAngle = PlrShoot(TurnOf).Angle + IIf(TurnOf = 1, 0, 180 - PlrShoot(TurnOf).Angle * 2)
        
        AnglBuf = PlrShoot(TurnOf).Angle 'set these to current values so if next time they are equal
        PwrBuf = PlrShoot(TurnOf).Power 'AIMER will not be redrawed
        
        DoEvents
        Sleep 20
    Loop Until GetAsyncKeyState(vbKeyReturn)
    
    
    TVisible = False
'>end -------------------------- AIMING

    Sleep 100
    GameLoop
    
    WaitSpace = True
End If

End Sub

Private Sub Form_Load()

Dim WinVer As OSVERSIONINFO

StClear

pTree1src.Picture = LoadResPicture(101, 0) 'set all the pictures and the
pTree2src.Picture = LoadResPicture(103, 0) 'masks in pictureboxes
pTree1msk.Picture = LoadResPicture(104, 0)
pTree2msk.Picture = LoadResPicture(102, 0)
picSrc.Picture = LoadResPicture(105, 0)
picMask.Picture = LoadResPicture(106, 0)
pExplmsk.Picture = LoadResPicture(108, 0)
pExplsrc.Picture = LoadResPicture(109, 0)
imgP1.Picture = LoadResPicture(104, 1)
imgP2.Picture = LoadResPicture(105, 1)

AllZero = True 'user will have to click Regenerate
InGame = False 'we aren't in game.
DispKeys = True 'display game keys this time

lsInfo.Caption = "Bang Bang Clone 32-bit. 2-player game in turns. Here you have to shoot at the other cannon by giving angle and power. The game generates mountains, hills and valleys as obstacles to your ball, just click 'Regenerate' to see. Click 'Start game' to play in the current terrain. Good luck!"
Me.Caption = Version
WinVer.dwOSVersionInfoSize = Len(WinVer)
X& = GetVersionEx(WinVer)
If WinVer.dwMajorVersion = 5 Then IsWinVer5 = True Else IsWinVer5 = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'if user "really want to quit", quit
If Not StopGame Then Cancel = 1

End Sub

Private Sub Form_Resize()
'it's legal to minimize the form...!
If Me.WindowState = vbMinimized Then Exit Sub
'if in game, don't allow the user to resize:
If InGame Then
    If PWinState = vbMaximized Then 'if previously maximized
        Me.WindowState = vbMaximized 're-set to maximized
        Exit Sub
    End If
    'if previously normal, re-set to normal
    If Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
    
    Me.Width = PrevW
    Me.Height = PrevH
    Exit Sub
End If

'Positionation, coloration and visualization stuff...
Back.Top = 15
Back.Left = 0
Back.Width = Me.ScaleWidth
Back.Height = Me.ScaleHeight - 30
lsInfo.Left = 5
lsInfo.Width = Back.Width - 10
lP1Score.Left = 8
lP1Score.Width = Me.ScaleWidth
lP2Score.Left = 0
lP2Score.Width = Me.ScaleWidth - 8
lP2Score.Top = Me.ScaleHeight - 15
cmdRegenerate.Top = Back.Height - 49
cmdStart.Top = Back.Height - 49
cmdRegenerate.Left = 24
cmdStart.Left = Back.Width - 121
AllZero = True
PrevH = Me.Height: PrevW = Me.Width
PWinState = Me.WindowState
lAct.Left = 0
lAct.Top = 10
lAct.Width = Back.ScaleWidth
'pWind.Left = Back.ScaleWidth / 2 - pWind.Width / 2
'pWind.Top = Back.ScaleHeight - 49
pWind.Left = Back.ScaleWidth - pWind.Width - 3
pWind.Top = Back.ScaleHeight - pWind.Height - 2
'lWind.Left = pWind.Left
'lWind.Top = Back.ScaleHeight - 23
lWind.Left = pWind.Left - lWind.Width - 5
lWind.Top = pWind.Top + pWind.Height / 2
Back.Cls
imgP1.Visible = False
imgP2.Visible = False

End Sub

Sub GameLoop()
'billions of declarations...
Const DeltaTime! = 0.01, BallMass! = 10, Friction! = 0.5, Grav! = 9.81
Dim CurTime!, CurAccelX!, CurAccelY!, PosX!, PosY!, SpeedX!, LastX!, SpeedY!, LastY!, I%, RetVal&
Dim DrawHole As Boolean, TerLvlBack!
'place a few DoEvents to ensure your code will work better (?)
DoEvents

PlrShoot(TurnOf).StartVX = IIf(TurnOf = 1, 1, -1) * Cos(PlrShoot(TurnOf).Angle * 3.141 / 180) * CInt(PlrShoot(TurnOf).Power) / 10
PlrShoot(TurnOf).StartVY = Sin(PlrShoot(TurnOf).Angle * 3.141 / 180) * CInt(PlrShoot(TurnOf).Power) / 10
CurTime = 0
LastY = Back.ScaleHeight - StartY(TurnOf)
LastX = StartX(TurnOf)
Back.AutoRedraw = False 'set this to false to speed up drawing, then set it back after

'BITBLT Rendering:
'picCopy is a temporary picture location
'all the pictureboxes must have autoredraw set to TRUE
'1) With lastX and lastY, copy previous background at lastX,lastY to the temporary location (picCopy)
RetVal = BitBlt(picCopy.hDC, 0, 0, 16, 16, Back.hDC, LastX - 8, Back.ScaleHeight - LastY - 8, SRCCOPY)
'2) AND the mask with the background at lastX,lastY
RetVal = BitBlt(Back.hDC, LastX - 8, Back.ScaleHeight - LastY - 8, 16, 16, picMask.hDC, 0, 0, SRCAND)
'3) INVERT with the source at lastX,lastY
RetVal = BitBlt(Back.hDC, LastX - 8, Back.ScaleHeight - LastY - 8, 16, 16, picSrc.hDC, 0, 0, SRCINVERT)

DoEvents
'calculate acceleration, then speed
CurAccelX = -Friction / BallMass * PlrShoot(TurnOf).StartVX
CurAccelY = -Friction / BallMass * PlrShoot(TurnOf).StartVY - Grav
SpeedX = PlrShoot(TurnOf).StartVX + Wind
SpeedY = PlrShoot(TurnOf).StartVY

WPlaySound "fire.wav" 'play a wonderful sound
'loop point
CycleRestart:
PosX = 1 / 2 * CurAccelX * DeltaTime ^ 2 + SpeedX + LastX
PosY = 1 / 2 * CurAccelY * DeltaTime ^ 2 + SpeedY + LastY

'here FIRST we copy the old background from picCopy to lastX,lastY
RetVal = BitBlt(Back.hDC, LastX - 8, Back.ScaleHeight - LastY - 8, 16, 16, picCopy.hDC, 0, 0, SRCCOPY)
'then we repeat as above, but with new X,Y coordinates
RetVal = BitBlt(picCopy.hDC, 0, 0, 16, 16, Back.hDC, PosX - 8, Back.ScaleHeight - PosY - 8, SRCCOPY)
RetVal = BitBlt(Back.hDC, PosX - 8, Back.ScaleHeight - PosY - 8, 16, 16, picMask.hDC, 0, 0, SRCAND)
RetVal = BitBlt(Back.hDC, PosX - 8, Back.ScaleHeight - PosY - 8, 16, 16, picSrc.hDC, 0, 0, SRCINVERT)

SpeedX = CurAccelX * DeltaTime + SpeedX + Wind
SpeedY = CurAccelY * DeltaTime + SpeedY
CurTime = CurTime + DeltaTime
LastX = PosX 'set new as old coords
LastY = PosY
CurAccelX = -Friction / BallMass * SpeedX
CurAccelY = -Friction / BallMass * SpeedY - Grav

'if ball is in extreme left or right, skip collision test as it will fall out
If (PosX > Back.ScaleWidth - 8) Or (PosX < 8) Then GoTo Collided

'if the ball is in cannon-zone, check if it has collided and if it is over the cannon
For I = 4 To 12
    If (PosX - 8 + I > P2BaseStart + 15 And PosX - 8 + I < P2BaseEnd - 15) Or (PosX - 8 + I > P1BaseStart + 15 And PosX - 8 + I < P1BaseEnd - 15) Then
        If Back.ScaleHeight - PosY + 8 >= TerLvl(PosX - 8 + I) Then

            If (PosX - 8 + I > P2BaseStart + 15 And PosX - 8 + I < P2BaseEnd - 15) Then CollidTo = 2
            If (PosX - 8 + I > P1BaseStart + 15 And PosX - 8 + I < P1BaseEnd - 15) Then CollidTo = 1
            GoTo Collided

        End If
    End If
Next I

'check if in at least one point the ball Y level is lower than the terrain level
For I = 1 To 16
    If Back.ScaleHeight - PosY + 6 >= TerLvl(PosX - 8 + I) Then

        CollidTo = 0
        GoTo Collided

    End If
Next I

Sleep 10

GoTo CycleRestart


Collided:
'set variables and play sound
If CollidTo = 2 And TurnOf = 1 Then HasWon = 1: WPlaySound "destroy.wav"
If CollidTo = 1 And TurnOf = 2 Then HasWon = 2: WPlaySound "destroy.wav"
If CollidTo = 2 And TurnOf = 2 Then HasWon = 1: WPlaySound "destroy.wav"
If CollidTo = 1 And TurnOf = 1 Then HasWon = 2: WPlaySound "destroy.wav"
If CollidTo = 0 Then HasWon = 0: WPlaySound "blnull.wav"

Sleep 100
'copy last old background on original position
RetVal = BitBlt(Back.hDC, LastX - 8, Back.ScaleHeight - LastY - 8, 16, 16, picCopy.hDC, 0, 0, SRCCOPY)
Back.AutoRedraw = True 'set to true now because it can cause complications

If HasWon = 0 Then GoTo NoWins 'if no one wins, skip this section
If CollidTo = 1 Then
    'draw the BOOM picture on the right location
    DrwTranspSpriteBlt Back, P1BaseStart - 4, TerLvl(P1BaseStart + 1) - 48, pExplsrc, pExplmsk
    DoEvents
    'and set the cannon picture to the "fired" one
    imgP1.Picture = LoadResPicture(107, 1)
Else
    'same things here
    DrwTranspSpriteBlt Back, P2BaseStart - 4, TerLvl(P2BaseStart + 1) - 48, pExplsrc, pExplmsk
    DoEvents
    imgP2.Picture = LoadResPicture(108, 1)
End If

'update score variable
PlrScore(HasWon) = PlrScore(HasWon) + 10 + IIf(10 - PlrShoots(HasWon) > 0, (10 - PlrShoots(HasWon)) * 8, 0)
'and label
lP1Score.Caption = "Player 1 score: " & PlrScore(1)
lP2Score.Caption = "Player 2 score: " & PlrScore(2)

DoEvents
Sleep 2500

MsgBox "Player " & HasWon & " wons!", vbInformation
CollidTo = 0
'ask if want to do another match
If MsgBox("Another match?", vbQuestion + vbYesNo) = vbYes Then
    CanShoot = False 'we can't shoot so if the user press space, nothing happens
    imgP1.Picture = LoadResPicture(104, 1) 'set "intact" cannon pictures
    imgP2.Picture = LoadResPicture(105, 1)
    ObjSet osLvSelect 'unhide all the LevelSelect controls
    cmdRegenerate_Click 'simulate a click to Regenerate
    Exit Sub
End If
'If the user want to stop game...
InGame = False
FrmMain.Back.Cls 'clear
StClear
VarResetFade 'create a new gradient
ObjSet osWelcome 'unhide all the Welcome control set
AllZero = True 'users has to click Regenerate
CanShoot = False 'user can't shoot
PlrScore(1) = 0 'reset score
PlrScore(2) = 0
cmdRegenerate_Click 'simulate a click in the Regenerate cmdbutton
Exit Sub

NoWins:
'if the shoot miss the target, draw a beautiful hole
DrawHole = True
'if the ball is in cannon-zone, don't draw the hole due to possible graphic
'corruption that I don't want to fix :)
For I = PosX - 4 To PosX + 20
    If (I > P1BaseStart And I < P1BaseEnd) Or (I > P2BaseStart And I < P2BaseEnd) Then DrawHole = False: Exit For
Next I

'if ball is in extreme left or right, don't draw the hole due to possible
'errors that I don't want to fix :)
If PosX < 8 Or PosX > Back.ScaleWidth - 8 Then DrawHole = False

'(but the first condition really helps!)

If DrawHole = True Then
    For I = PosX - 12 To PosX + 12
        If Not (I > P1BaseStart And I < P1BaseEnd) Or (I > P2BaseStart And I < P2BaseEnd) Then

            TerLvlBack = TerLvl(I) 'backup precedent terrain level
            'I use a parabolic shape to draw the hole
            'don't ask me to explain the formula. I don't remember!!!;)
            TerLvl(I) = Int(TerLvl(I) + (36 - (((I - (PosX)) / 2) ^ 2)))
            Back.Line (I, TerLvlBack)-(I, TerLvl(I)), RGB(134, 69, 0)
            Back.Line (I - 1, TerLvl(I - 1))-(I, TerLvl(I)), RGB(0, 0, 0)

        End If
    Next I
End If

PlrShoots(TurnOf) = PlrShoots(TurnOf) + 1 'update player shoot number
If TurnOf = 1 Then TurnOf = 2 Else TurnOf = 1 'switch turn
lAct.Visible = True 'show the PRESS SPACE label

DoEvents

End Sub

Sub VarResetFade()

FStep(1) = 0
FStep(2) = Back.ScaleHeight
Randomize (Timer)
R(1) = Int(Rnd * 255) 'randomize colors
G(1) = Int(Rnd * 255)
B(1) = Int(Rnd * 255)
R(2) = Int(Rnd * 255)
G(2) = Int(Rnd * 255)
B(2) = Int(Rnd * 255)

ObjFade Back, drVertical

End Sub

Sub ObjSet(Status As Byte)

'these are all instruction to position, color and hide/unhide
'object needed to the game phases: Welcome, Playing, LevelSelect
Select Case Status
Case osPlaying
    lsInfo.Visible = False
    cmdStart.Visible = False
    cmdRegenerate.Visible = False
    lP1Score.Caption = "Player 1 score: " & PlrScore(1)
    lP2Score.Caption = "Player 2 score: " & PlrScore(2)
    lAct.Visible = True
    pWind.Visible = True
    lWind.Visible = True
    pWind.Cls

Case osLvSelect
    lsInfo.Caption = "Click Regenerate until you see a terrain you like and then click Start Game to start another game."
    lsInfo.Visible = True
    cmdRegenerate.Visible = True
    cmdStart.Visible = True
    lWind.Visible = False
    pWind.Visible = False
    imgP1.Picture = LoadResPicture(104, 1)
    imgP2.Picture = LoadResPicture(105, 1)
    TVisible = False

Case osWelcome
    lsInfo.Visible = True
    cmdStart.Visible = True
    cmdRegenerate.Visible = True
    lAct.Visible = False
    imgP1.Visible = False
    imgP2.Visible = False
    lsInfo.Caption = "Bang Bang Clone 32-bit. 2-player game in turns. Here you have to shoot at the other cannon by giving angle and power. The game generates mountains, hills and valleys as obstacles to your ball, just click 'Regenerate' to see. Click 'Start game' to play in the current terrain. Good luck!"
    lP1Score.Caption = ""
    lP2Score.Caption = ""
    pWind.Visible = False
    lWind.Visible = False
    prWind(1).Visible = False
    prWind(2).Visible = False
    imgP1.Picture = LoadResPicture(104, 1)
    imgP2.Picture = LoadResPicture(105, 1)
    TVisible = False
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)

frmEnd.Show 1, Me

End 'I have to stop the program in this mode, otherwise if I try to exit
'when I'm AIMING, it really doesn't exit, just keep looping in the Do..Loop
End Sub

Property Let AimX(CX As Integer)
'positioning stuff
AimTgt.Left = CX - AimTgt.Width / 2
TgtLX.X1 = CX - 14
TgtLX.X2 = CX + 14
TgtLY.X1 = CX
TgtLY.X2 = CX

End Property

Property Get AimX() As Integer

AimX = AimTgt.Left + AimTgt.Width / 2

End Property

Property Let AimY(CY As Integer)

AimTgt.Top = CY - AimTgt.Height / 2
TgtLY.Y1 = CY - 14
TgtLY.Y2 = CY + 14
TgtLX.Y1 = CY
TgtLX.Y2 = CY

End Property

Property Get AimY() As Integer

AimY = AimTgt.Top + AimTgt.Height / 2

End Property

Property Let TVisible(ans As Boolean)
'hiding/unhiding stuff
If ans Then

    AimTgt.Visible = True
    AimLine.Visible = True
    TgtLX.Visible = True
    TgtLY.Visible = True

Else
    
    AimTgt.Visible = False
    AimLine.Visible = False
    TgtLX.Visible = False
    TgtLY.Visible = False
End If

End Property

Property Get TVisible() As Boolean

If AimTgt.Visible Then TVisible = True Else TVisible = False

End Property

Property Let TgtAngle(TAngl As Integer)

Dim SinY!, CosX!

SinY = Sin(TAngl * 3.14 / 180) 'calculate how much the AIMER must be wide
CosX = Cos(TAngl * 3.14 / 180) 'and high
AimX = AimStartX + CosX * PlrShoot(TurnOf).Power
AimY = AimStartY - SinY * PlrShoot(TurnOf).Power
AimLine.X2 = AimX
AimLine.Y2 = AimY

End Property

Property Let AimStartX(CX As Integer)
AimLine.X1 = CX
End Property

Property Get AimStartX() As Integer
AimStartX = AimLine.X1
End Property

Property Let AimStartY(CY As Integer)
AimLine.Y1 = CY
End Property

Property Get AimStartY() As Integer
AimStartY = AimLine.Y1
End Property

Property Let AimerColor(NewColor As Long)
'coloring stuff
AimLine.BorderColor = NewColor
AimTgt.BorderColor = NewColor
TgtLX.BorderColor = NewColor
TgtLY.BorderColor = NewColor

End Property

Property Get AimerColor() As Long
AimerColor = AimLine.BorderColor
End Property
