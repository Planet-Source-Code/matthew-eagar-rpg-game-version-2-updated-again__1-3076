Attribute VB_Name = "modGame"
'The BitBlt function allows for fast and smooth drawing to the form
'and to picture boxes, but isn't as fast as it should be for making games.

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal animX As Long, ByVal animY As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'allows the playing of wav files
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'for the sound function
Public Const SND_SYNC = &H0         '  play synchronously
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound

'for the bitblt function
Public Const SRCCOPY = &HCC0020   'Copies the source over the destination
Public Const SRCINVERT = &H660046 'Copies and inverts the source over the destination
Public Const SRCAND = &H8800C6    'Adds the source to the destination
Public Const oSRCERASE = &H440328
Public Const oPATPAINT = &HFB0A09
Public Walkable(0 To 899) As Integer
Public Texture(0 To 899) As Integer
Public tileLeft(0 To 899) As Integer
Public tileTop(0 To 899) As Integer

