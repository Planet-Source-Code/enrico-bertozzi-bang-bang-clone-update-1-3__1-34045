Attribute VB_Name = "BB32CModule"
'HOW CAN I SET THE LINE STARTING COORDINATES? AND THE COLOR?
'Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Playsound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long 'filename,0,1
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Global Const Version$ = "Bang Bang Clone! version 1.3g"

Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCINVERT = &H660046
Public Const SRCAND = &H8800C6

Public Const osPlaying As Byte = 0
Public Const osLvSelect As Byte = 1
Public Const osWelcome As Byte = 2

Global PrevWind(1 To 2) As Single
Global TerLvl(2000) As Integer
Global TurnOf As Byte, HasWon As Byte, CollidTo As Byte
Global PlrShoot(1 To 2) As ShootCfg
Global PlrScore(1 To 2) As Integer
Global PlrShoots(1 To 2) As Integer
Global CanShoot As Boolean
Global WaitSpace As Boolean
Global DispKeys As Boolean
Global Wind As Single
Global StartX(1 To 2) As Integer
Global StartY(1 To 2) As Integer
Global InGame As Boolean, PrevH As Integer, PrevW As Integer, PWinState As Byte, AllZero As Boolean
Global P1BaseStart%, P1BaseEnd%, P2BaseStart%, P2BaseEnd%

Global IsWinVer5 As Boolean

Public Type ShootCfg

    StartVX As Single
    StartVY As Single
    Angle As Byte
    Power As Byte

End Type

Public Type OSVERSIONINFO

   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128

End Type

Public Function StopGame() As Boolean
'ask if the user want to stop the game
If Not InGame Then
    StopGame = True
Else
    If MsgBox("Really want to end game?", vbApplicationModal + vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then StopGame = True Else StopGame = False
End If

End Function

Public Sub WPlaySound(FileName As String)

Dim S&
S = Playsound(FileName, 0, 1)

End Sub

Public Sub StClear()
'this clear all the color variables for the gradient generator
For I = 0 To 21
    R(I) = 0
    G(I) = 0
    B(I) = 0
    FStep(I) = 0
Next I

End Sub

Public Function DrwTranspSpriteBlt(pTo As Object, pToX As Integer, pToY As Integer, pFrom As Object, pMask As Object)

Static E&
E = BitBlt(pTo.hDC, pToX, pToY, pFrom.Width, pFrom.Height, pMask.hDC, 0, 0, SRCAND)
DrwTranspSpriteBlt = E And BitBlt(pTo.hDC, pToX, pToY, pFrom.Width, pFrom.Height, pFrom.hDC, 0, 0, SRCINVERT)

End Function

'Public Function DrwSpriteBlt(pTo As Object, pToX As Integer, pToY As Integer, pFrom As Object)
'(NOT USED)
'Static E&
'E = BitBlt(pTo.hDC, pToX, pToY, pFrom.Width, pFrom.Height, pFrom.hDC, 0, 0, SRCCOPY)
'
'End Function
