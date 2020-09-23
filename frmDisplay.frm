VERSION 5.00
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8595
   ClientLeft      =   150
   ClientTop       =   1050
   ClientWidth     =   10365
   ControlBox      =   0   'False
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmDisplay.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   8595
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   23595
      Left            =   -22800
      ScaleHeight     =   23535
      ScaleWidth      =   27420
      TabIndex        =   2
      Top             =   -19320
      Visible         =   0   'False
      Width           =   27480
   End
   Begin VB.PictureBox picBuf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   -3840
      ScaleHeight     =   6735
      ScaleWidth      =   8850
      TabIndex        =   1
      Top             =   -2040
      Width           =   8850
   End
   Begin VB.PictureBox picSpr 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6195
      Left            =   -360
      Picture         =   "frmDisplay.frx":0A1C
      ScaleHeight     =   6135
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   -600
      Visible         =   0   'False
      Width           =   6195
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'========================================================='
'========================================================='
'=============== RPG Game Version 0.0.2 =================='
'============== Written by Matthew Eagar ================='
'============ Compiled in Visual Basic 6.0 ==============='
'========================================================='
'========================================================='
'
'   This program is an example of an RPG game engine made
'   with VB 6.0.  I drew all the graphics in MS Paint,
'   and all coding is origional.
'
'   This isn't ment to be a full game, just a working engine.
'   there is no actual objective.  I havn't yet got doors
'   working, because that would require me to draw some more
'   textures for the insides of houses, which takes FOREVER!
'   Also, the textures could REALLY use some work,
'   as they were drawn in MS Paint.
'
'   This program may not run well on some computers.
'   The method used, bitblt, works well, but isn't designed for games.
'   It runs fine on a Pentium 233, but slow on a P75.  I havn't tested
'   it on anything in between those.
'
'   I'm still working on this, so look for me to post newer versions
'   of it.  It'll remain free, and it's really ment for educational purposes.
'
'   Please contact me with ANY questions, comments, suggestions, or problems,
'   ANY input is welcome:
'
'   email:  meagar@home.com
'   ICQ:    45058462
'
'   Also, I havn't tested this on any computer running anything less then VB6.
'   I did run it in vb5, but it took some work.
'   You will need the VB6 runtime files the use this.
'
'   Updates to Version 2:
'   Added side scrolling and top scrolling
'   Rechanged the map size from 13x11 to 30x30 tiles to accomidate side scrolling
'   Added Bridge Tiles for bridge construction
'   Added sound effects
'   re-wrote most of movement code
'


Dim animX As Integer    'holds the current x location of the animation frame
Dim animY As Integer    'holds the current y location of the animation frame

Dim direction As Integer    'the direction the characters facing
Dim charX As Integer       'holds the character's x coords
Dim charY As Integer       'holds the character's y coords
Dim lastX As Integer    'holds the character's last y coords
Dim lastY As Integer    'holds the character's last x coords
Dim BackBuilt As Integer 'determines if the back ground needs to be built
Dim Speed As Integer    'holds the current speed, set by pressing the + or - keys
Dim mapx As Integer     'holds the current map x number
Dim mapy As Integer     'holds the current map y number
Dim MapName As String   'holds the name of the map
Dim screenX As Integer  'holds the current location of the screen on the map
Dim screenY As Integer  'holds the current location of the screen on the map
Dim charPosX As Integer 'holds the coords to center the character on the screen
Dim charPosY As Integer 'holds the coords to center the character on the screen
Dim sound As Boolean     'holds whether to play sounds or not
Dim moving As Integer
Dim changeFrame As Integer

'symbolic constants
'directions
Const dLEFT As Integer = 1    'left direction
Const dUP As Integer = 2      'up direction
Const dRIGHT As Integer = 3   'right direction
Const dDOWN  As Integer = 4   'down direction

'animation frames
Const aLEFT As Integer = 2    'left animation
Const aUP As Integer = 104    'up animation
Const aRIGHT As Integer = 206 'right animation
Const aDOWN As Integer = 308  'down animation


'when the user presses a key
Private Sub picBuf_KeyDown(KeyCode As Integer, Shift As Integer)

Dim X As Integer 'counting variable

'if movement, turn the mouse cursor into the invisible icon.
'simply making a mouse cursor that was invisible is easier
'then using API calls.
frmDisplay.MouseIcon = frmTextures.picInvisible.Picture

'determine how to act, based on which key the user presses.
Select Case KeyCode
Case Is = 37    'left arrow key
    animX = aLEFT   'set the animation frame to the proper direction
    direction = dLEFT 'set the direction
Case Is = 38    'up arrow key
    animX = aUP 'set the animation frame to the proper direction
    direction = dUP
Case Is = 39    'right arrow key
    animX = aRIGHT
    direction = dRIGHT
Case Is = 40    'down arrow key
    animX = aDOWN
    direction = dDOWN
Case Is = 27    'escape key
    'ask if the user would like to exit
    If MsgBox("Are you sure you would like to exit?", vbYesNo + vbDefaultButton2 + vbQuestion, "Exit?") = vbYes Then End
Case Is = 109   'minus key - increases screen size
    Speed = Speed - 2
    If Speed < 5 Then Speed = 5
Case Is = 107   'plus key - decreases screen size
    Speed = Speed + 2
    If Speed > 30 Then Speed = 30
Case Is = 83    'the S key
    'turn sound on or off
    If sound = True Then
        sound = False
    Else
        sound = True
    End If
End Select

'see if the movement timer should be enabled
If KeyCode >= 37 And KeyCode <= 40 Then  'if a direction key's been pressed
    moving = 1
    Call moveChar
End If
End Sub

Private Sub picBuf_KeyUp(KeyCode As Integer, Shift As Integer)

moving = 0
End Sub

Private Sub Form_Load()

'initialize the variables
animX = 2
animY = 1
screenX = 10
screenY = 10
charX = screenX + charPosX + 25
charY = screenY + charPosY + 25
sound = True
BackBuilt = False

'maps are loaded in the following way:
'take the mapX, then add the letter 'a' then take the mapY, then add ".map"
'so, the first map is called 0a0.map, the map beside it is called
'1a0.map, and the map above the first is called 0a1.map
'eventually the middle letter will stand for the area, eg a = lev 1, b = lev 2

mapx = 0    'the current map
mapy = 0

Speed = 15  'set the initial walking speed

'set the size of the main picture box
'change the number to make the picture bigger or smaller, but the number can't be
'large then 1 or smaller then 0
picBuf.Height = Int(Screen.Height * 0.85)
picBuf.Width = Int(Screen.Width * 0.85)

picBuf.Left = (Screen.Width - picBuf.Width) / 2 'center the main picture box
picBuf.Top = (Screen.Height - picBuf.Height) / 2

'charPosX is the distance of the character from the left side of the screen
'charPosY is the distance from the top
charPosX = picBuf.Width * 0.03  'center the character on the screen
charPosY = picBuf.Height * 0.03

Call BuildBack  'build the back ground
Call redrawPic  'load the pic into the main pic box

Call moveChar
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, animX As Single, animY As Single)
'turn the mouse icon into the visible icon
frmDisplay.MouseIcon = frmTextures.picVisible.Picture
End Sub

'assembles the back ground
Sub BuildBack()

'this sub builds the back ground.  It is called only once per map,
'as the map is built in a hidden pic box, and kept untill the next map is needed.

Dim g As Integer    'counting variable
Dim a As Integer    'temp variable
Dim X As Integer    'holds x coords of tile
Dim Y As Integer    'holds y coords of tile
On Error GoTo errHandler

'set the name of the map
If Right(App.Path, 1) = "\" Then
    MapName = App.Path & mapx & "a" & mapy & ".map"
Else
    MapName = App.Path & "\" & mapx & "a" & mapy & ".map"
End If

'read the textures and the walkable values from the map file
Open MapName For Input As #1
    For g = 0 To 899
        Input #1, Texture(g), Walkable(g)
    Next g
Close

'clear the picture box which will hold the back ground
picBack.Cls

X = 0
Y = 0

'loop through each tile, getting it with bitblt from frmTextures, and putting it into
'the picBack pic box.
For g = 0 To 899
    tileLeft(g) = X
    tileTop(g) = Y
    a = BitBlt(picBack.hDC, X, Y, 40, 40, frmTextures.picTextures(Texture(g)).hDC, 0, 0, SRCCOPY)
    Y = Y + 40
    
    'if a column has been finished, goto the next one
    If Y >= 1200 Then
        Y = 0
        X = X + 40
    End If
Next g

'by-pass error handler
GoTo endsub

'for errors
errHandler:

MsgBox "Error number " & Err.Number & ": " & Err.Description, vbOKOnly + vbCritical, "Dragon Lore"
MsgBox MapName & " was not found or was corrupted.  Please re-install this program."
End

endsub:
End Sub

Sub redrawPic()

'this function draws the picture to the screen.

'black out the old picture
picBuf.Cls
'Copy the back ground to the buffer pic box
Call BitBlt(picBuf.hDC, 0, 0, 2900, 9500, picBack.hDC, screenX, screenY, SRCCOPY)
'Copy the first layer of the sprite to the buffer
'this mask is like a negative, it is a black shadow of the character,
'sourouneded by white (see picSpr). when added using SRCAND, every other color
'except black becomes transparent.  So only the black figure is left
Call BitBlt(picBuf.hDC, charPosX, charPosY, 50, 50, picSpr.hDC, animX + 50, animY, SRCAND)
'Copy the second layer of the sprite to the buffer, for transparent effect.
'this copys the color onto the black, using SRCINVERT, which makes it like copying
'the colors onto white, so the colors stay the same.
Call BitBlt(picBuf.hDC, charPosX, charPosY, 50, 50, picSpr.hDC, animX, animY, SRCINVERT)
'refresh the picture
picBuf.Refresh

End Sub

Private Function touching() As Integer
Dim g As Integer ' counting variable
Dim tmpX As Integer
Dim tmpY As Integer


'this looks at the direction the character is moving, and sees if the next step
'will put the character onto a tile which has a walkable value of 1, which is
'either water trees or a building.  If it is, it returns 1. if not, it returns 0.

tmpX = 0
tmpY = 0

'check each tile
'I'm looking for ways to OPTIMIZE this!! Email me with suggestions!
For g = 0 To 899
    
    'only proceed to check a tile if it is within a certain radius of the character,
    'and if it is a tree/water/wall
    If Abs((charX + 25) - (tileLeft(g) + 20)) < 250 And Abs((charY + 25) - (tileTop(g) + 20)) < 250 And Walkable(g) = 1 Then
        If direction = dLEFT Then   'if the character is walking left
            'check the left side of the character
            If charX - 25 - Speed > tileLeft(g) And charX - 25 - Speed < tileLeft(g) + 40 Then
                'check the lower left corner
                If charY + 25 > tileTop(g) And charY + 25 < tileTop(g) + 40 Then
                    GoTo endsub
                'check the top left corner
                ElseIf charY - 25 > tileTop(g) And charY - 25 < tileTop(g) + 40 Then
                    GoTo endsub
                'check the center of the left side
                ElseIf charY > tileTop(g) And charY < tileTop(g) + 40 Then
                    GoTo endsub
                End If
            End If
        ElseIf direction = dUP Then 'if the character is walking up
            'check the top side of the character
            If charY - 25 - Speed > tileTop(g) And charY - 25 - Speed < tileTop(g) + 40 Then
                'check the top right corner
                If charX + 25 > tileLeft(g) And charX + 25 < tileLeft(g) + 40 Then
                    GoTo endsub
                'check to top left corner
                ElseIf charX - 25 > tileLeft(g) And charX - 25 < tileLeft(g) + 40 Then
                    GoTo endsub
                'check the center of the top side
                ElseIf charX > tileLeft(g) And charX < tileLeft(g) + 40 Then
                    GoTo endsub
                End If
            End If
        ElseIf direction = dRIGHT Then  'if the character is walking right
            'check the right side of the character
            If charX + 25 + Speed > tileLeft(g) And charX + 25 + Speed < tileLeft(g) + 40 Then
                'check the right top corner
                If charY - 25 > tileTop(g) And charY - 25 < tileTop(g) + 40 Then
                    GoTo endsub
                ElseIf charY + 25 > tileTop(g) And charY + 25 < tileTop(g) + 40 Then
                    GoTo endsub
                'check the center of the right side
                ElseIf charY > tileTop(g) And charY < tileTop(g) + 40 Then
                    GoTo endsub
                End If
            End If
        ElseIf direction = dDOWN Then   'if the character is walking down
            'check the bottom side of the character
            If charY + 25 + Speed > tileTop(g) And charY + 25 + Speed < tileTop(g) + 40 Then
                'check the bottom right corner
                If charX + 25 > tileLeft(g) And charX + 25 < tileLeft(g) + 40 Then
                    GoTo endsub
                ElseIf charX - 25 > tileLeft(g) And charX - 25 < tileLeft(g) + 40 Then
                    GoTo endsub
                'check the middle of the bottom side
                ElseIf charX > tileLeft(g) And charX < tileLeft(g) + 40 Then
                    GoTo endsub
                End If
            End If
        End If
    End If
Next g

touching = 0

GoTo endFunct

endsub:

'reset the character location
touching = 1

endFunct:

End Function

'move the character
Sub moveChar()
Dim lastlastX, lastlastY As Integer

'copy the current location of the character into the lastx and lasty variables.
  
           
While (moving = 1)  'as long as a key is pressed

    If touching() <> 1 Then
        'move the character in the proper direction
        If direction = dLEFT Then
            screenX = screenX - Speed
        ElseIf direction = dUP Then
            screenY = screenY - Speed
        ElseIf direction = dRIGHT Then
            screenX = screenX + Speed
        ElseIf direction = dDOWN Then
            screenY = screenY + Speed
        End If
                
        charX = screenX + charPosX + 25
        charY = screenY + charPosY + 25
            
        Call redrawPic 'redraws the form
        
        'this causes the frame to be updated once every 2 loops
        If changeFrame = 1 Then
            animY = animY + 51    'advance the frame, each frame is 50 pixels wide, + a 1 pixel border
            changeFrame = 0
                
            'there are 8 frames in the character's animation: this sees if the last frame has
            'been shown. if it has, it resets it to the first.
            If animY >= 408 Then
                animY = 1  'goes to first frame
                If sound = True Then Call sndPlaySound(App.Path & "\" & "1.wav", SND_ASYNC) 'play the foot step sound
            ElseIf animY >= 204 And animY <= 255 Then
                If sound = True Then Call sndPlaySound(App.Path & "\" & "1.wav", SND_ASYNC)  'play the foot step sound
            End If
        Else
            changeFrame = 1
        End If
    Else
        'if it is thouching a non-walkable tile like trees/water/buildings then
        moving = 0
        Exit Sub
    End If
    
    'see if the back ground has been built
    If BackBuilt = False Then
        'build the background
        Call BuildBack
        BackBuilt = True
    End If
    
    
    'see if the character has left the screen, by checking if the character's
    'x or y position is greater then the total amount of tiles
    If screenX + 25 >= 1200 - charPosX Then 'if the character has left the right side of the screen
        mapx = mapx + 1 'set the current map name to the next map name
        screenX = 10 - charPosX - 25 'set the character's position back to the left side of the screen
        Call BuildBack  'redraw the back ground

    ElseIf screenX + 25 <= 0 - charPosX Then 'see if the character has left the left side of the screen
        mapx = mapx - 1 'set the current map name to the next map name
        screenX = 1190 - charPosX - 25 'set the character position to the right side of the screen
        Call BuildBack  'redraw the back ground

    ElseIf screenY + 25 <= 0 - charPosY Then  'see if the character has left the top of the screen
        mapy = mapy + 1 'set the current map name to the next map name
        screenY = 1190 - charPosY - 25 'set the characters position to the bottom of the screen
        Call BuildBack  'redraw the back ground

    ElseIf screenY + 25 >= 1200 - charPosY Then 'see if the character has left the bottom of the screen
        mapy = mapy - 1 'set the current map name to the next map name
        screenY = 10 - charPosY - 25  'move the character to the top of the screen
        Call BuildBack  'redraw the back ground
    End If

    DoEvents    'allow for keyup
Wend

End Sub

