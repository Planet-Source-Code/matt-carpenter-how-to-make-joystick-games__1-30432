VERSION 5.00
Begin VB.Form frmJoyStick 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   365
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmJoyStick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright Matt Carpenter- 2002
Option Explicit
Private Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Private Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long

Const MAXPNAMELEN = 32

Private Type JOYCAPS
        wMid As Integer
        wPid As Integer
        szPname As String * MAXPNAMELEN
        wXmin As Long
        wXmax As Long
        wYmin As Long
        wYmax As Long
        wZmin As Long
        wZmax As Long
        wNumButtons As Long
        wPeriodMin As Long
        wPeriodMax As Long
End Type

Private Type JOYINFO
        wXpos As Long
        wYpos As Long
        wZpos As Long
        wButtons As Long
End Type

'Joystick error codes and return values
Const JOYERR_NOERROR = 0
Const JOYERR_BASE As Long = 160
Const JOYERR_UNPLUGGED As Long = (JOYERR_BASE + 7)
Const MMSYSERR_BASE As Long = 0
Const MMSYSERR_NODRIVER As Long = (MMSYSERR_BASE + 6)
Const MMSYSERR_INVALPARAM As Long = (MMSYSERR_BASE + 11)
Const JOYSTICK1 As Long = &H0
Const JOYSTICK2 As Long = &H1
Const JOY_BUTTON2 = &H2
Const JOY_BUTTON1 = &H1



Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000


'Sprite and Mask containers
Dim DCMask As Long
Dim DCSpriteBlack As Long
Dim DCSpriteRed As Long
Dim DCSpriteInner As Long

'Flag for ending the game loop
Dim TimeToEnd As Boolean

'Variable to hold the Max Y and Max X values
Dim MaxX As Long
Dim MaxY As Long
'The minimum values
Dim MinX As Long
Dim MinY As Long

'The relation value sbetween the Joystick position and the
'relative window position
Dim RelativeX As Long
Dim RelativeY As Long


Const SpriteWidth As Long = 32
Const SpriteHeight As Long = 32

Const HalfSpriteWidth As Long = SpriteWidth / 2
Const HalfSpriteHeight As Long = SpriteHeight / 2


'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

Private Sub Form_Load()
Dim rt As Long
Dim JoyTestInfo As JOYINFO
Dim JoyStickCaps As JOYCAPS


rt = joyGetPos(JOYSTICK1, JoyTestInfo) 'See if there is a joystick

If rt <> JOYERR_NOERROR Then
    If rt = JOYERR_UNPLUGGED Then
        MsgBox "No joystick connected" & vbCrLf & "Finishing..."
    ElseIf rt = MMSYSERR_NODRIVER Then
        MsgBox "No Joystick driver on system" & vbCrLf & "Finishing..."
    Else
        MsgBox "Unknown Error" & vbCrLf & "finishing..."
    End If
        
    Unload Me
    Exit Sub
End If

'Get the max and min position on the joystick
joyGetDevCaps JOYSTICK1, JoyStickCaps, Len(JoyStickCaps)

With JoyStickCaps

    MaxX = .wXmax
    MinX = .wXmin
    MaxY = .wYmax
    MinY = .wYmin
    
End With

'Load the images
DCSpriteBlack = GenerateDC(App.Path & "\crossblack.bmp")
DCSpriteInner = GenerateDC(App.Path & "\crossinner.bmp")
DCSpriteRed = GenerateDC(App.Path & "\crossred.bmp")
DCMask = GenerateDC(App.Path & "\crossm.bmp")

RunMainGame


End Sub

Private Sub Form_Resize()


RelativeX = MaxX / Me.ScaleWidth

RelativeY = MaxY / Me.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

TimeToEnd = True

End Sub


Private Sub RunMainGame()
Dim X As Long, Y As Long
Dim JoyInformation As JOYINFO

Me.Show

Do
DoEvents


Me.Cls

joyGetPos JOYSTICK1, JoyInformation

X = (JoyInformation.wXpos / RelativeX) - HalfSpriteWidth
Y = (JoyInformation.wYpos / RelativeY) - HalfSpriteHeight
'draw sprites
BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, DCMask, 0, 0, vbSrcAnd
BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, DCSpriteBlack, 0, 0, vbSrcPaint
'Determine if any buttons are pressed and draw the right images to be showed
If (JoyInformation.wButtons And JOY_BUTTON1) Then
    BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, DCSpriteRed, 0, 0, vbSrcPaint
End If

If (JoyInformation.wButtons And JOY_BUTTON2) Then
    BitBlt Me.hdc, X, Y, SpriteWidth, SpriteHeight, DCSpriteInner, 0, 0, vbSrcPaint
End If

'Make it appear
Me.Refresh

DoEvents
Loop Until TimeToEnd


DeleteGeneratedDC DCMask
DeleteGeneratedDC DCSpriteBlack
DeleteGeneratedDC DCSpriteRed
DeleteGeneratedDC DCSpriteInner

End Sub


