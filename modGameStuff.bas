Attribute VB_Name = "modGameStuff"
' Joystick stuff
Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON3 = &H4
Public Const JOY_BUTTON4 = &H8
Public Const JOYERR_BASE = 160
Public Const JOYERR_NOERROR = (0)  '  no error
Public Const JOYERR_NOCANDO = (JOYERR_BASE + 6) '  request not completed
Public Const JOYERR_PARMS = (JOYERR_BASE + 5) '  bad parameters
Public Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7) '  joystick is unplugged
Public Const JOYSTICKID1 = 0
Public Const JOYSTICKID2 = 1
Public Const MAXPNAMELEN = 32  '  max product name length (including NULL)

Type JOYCAPS
        wMid As Integer
        wPid As Integer
        szPname As String * MAXPNAMELEN
        wXmin As Integer
        wXmax As Integer
        wYmin As Integer
        wYmax As Integer
        wZmin As Integer
        wZmax As Integer
        wNumButtons As Integer
        wPeriodMin As Integer
        wPeriodMax As Integer
End Type
Global MyJoyCaps As JOYCAPS
Type JOYINFO
        wXpos As Integer
        wYpos As Integer
        wZpos As Integer
        wButtons As Integer
End Type
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long

' Graphics stuff
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
' BitBlt()'s parameters
'1. Destination Handle
'2. Destination X
'3. Destination Y
'4. Width of graphic being painted
'5. Height of graphic being painted
'6. Source Handle
'7. Source X
'8. Source Y
'9. Raster Operation
' Miscellanous
Type tBitmap
     Top As Long
     Left As Long
     Right As Long
     Bottom As Long
     Height As Long
     Width As Long
End Type

'Variable of DTG game
'Public Increment As Integer

