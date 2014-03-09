VERSION 5.00
Begin VB.Form frmGeometryWars 
   BackColor       =   &H00000000&
   Caption         =   "Geometry Wars"
   ClientHeight    =   3090
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrKeyboardInput 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   2760
   End
   Begin VB.Timer tmrMoveBullets 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   2760
   End
   Begin VB.Timer tmrMoveShapes 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   2760
   End
   Begin VB.Timer tmrSpawnShapes 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   600
      Top             =   2760
   End
   Begin VB.Timer tmrMoveShip 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblLevel 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Level: 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblScoreMultiplier 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Score multiplier: 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3495
   End
   Begin VB.Image imgExplosion 
      Height          =   735
      Left            =   5400
      Picture         =   "frmGeometryWars.frx":0000
      Stretch         =   -1  'True
      Top             =   2040
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Line linShoot 
      BorderColor     =   &H80000005&
      X1              =   840
      X2              =   840
      Y1              =   2520
      Y2              =   2160
   End
   Begin VB.Line linOutline 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Index           =   3
      X1              =   1800
      X2              =   2520
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line linOutline 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Index           =   4
      X1              =   2640
      X2              =   3360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line linOutline 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Index           =   2
      X1              =   960
      X2              =   1680
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line linOutline 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Index           =   1
      X1              =   120
      X2              =   840
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label lblScore 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Score: 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image imgShip 
      Height          =   390
      Left            =   600
      Picture         =   "frmGeometryWars.frx":3EA4
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label lblBombs 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   ": 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLives 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   ": 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Image imgBombs 
      Height          =   315
      Left            =   7920
      Picture         =   "frmGeometryWars.frx":48A6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   300
   End
   Begin VB.Image imgLives 
      Height          =   300
      Left            =   7200
      Picture         =   "frmGeometryWars.frx":4ABC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   300
   End
   Begin VB.Image imgShapeImages 
      Height          =   375
      Index           =   4
      Left            =   4800
      Picture         =   "frmGeometryWars.frx":4CAE
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgShapeImages 
      Height          =   690
      Index           =   3
      Left            =   3960
      Picture         =   "frmGeometryWars.frx":545C
      Top             =   2040
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgShapeImages 
      Height          =   690
      Index           =   2
      Left            =   3120
      Picture         =   "frmGeometryWars.frx":6DC6
      Top             =   2040
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgShapeImages 
      Height          =   705
      Index           =   1
      Left            =   2280
      Picture         =   "frmGeometryWars.frx":8730
      Top             =   2040
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgShapeImages 
      Height          =   525
      Index           =   0
      Left            =   1680
      Picture         =   "frmGeometryWars.frx":A06A
      Top             =   2160
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imgShape 
      Height          =   495
      Index           =   1
      Left            =   5400
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgBullet 
      Height          =   195
      Index           =   1
      Left            =   1200
      Picture         =   "frmGeometryWars.frx":AF70
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmGeometryWars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name: Michael Parrott
'Due date: Wednesday, June 16, 2010
'Course Code: ICS201
'Teacher: Mr. Fernandes
'Title: Final Project - Geometry Wars
'Description: 2D top down space shooter
'Major Skills: Enumerations
'Trigonometric functions
'Dynamic arrays
'Declared functions
'Constants
'Added Features: Multiple Shapes
'Bombs
'Notes: View at a screen resolution of 1024x768.
'Everything is centred at 1024x768
Option Explicit

'Declared functions
Private Declare Function GetTickCount Lib "kernel32" () As Long 'For millisecond timing

'Enumerations
Private Enum ShapeType
    'Different types of shapes that the user will have to shoot
    'Used mainly for determining the AI
    WandererShape
    GruntShape
    WeaverShape
    SpinnerShape
    TinySpinnershape
End Enum

Private Enum GunType
    'Can be used for power ups and changing how the ship shoots
    DefaultGun
End Enum

'User defined types
Private Type Shape
    X As Long 'X position
    Y As Long 'Y position
    xVel As Integer
    yVel As Integer
    type As ShapeType 'The type of the shape
    controlIndex As Integer
End Type

Private Type Ship 'What the user will be using
    X As Single 'Co-ordinates
    Y As Single
    xVel As Single 'Velocity
    yVel As Single
    gun As GunType 'What type of gun the ship has
    rotation As Integer 'From 0-359 degrees, where 0 degrees is pointing up ^
End Type

Private Type Bullet
    X As Integer 'Co-ordinates
    Y As Integer
    xVel As Integer 'Velocity
    yVel As Integer
    controlIndex As Integer
End Type

'Constants
Const PLAYER_SPEED As Integer = 3 'How much the player's speed changes when they press the arrow keys
Const MAX_PLAYER_SPEED As Integer = 75 'The faster the player can move
Const LINE_LENGTH As Integer = 500 'The length of the line to indicate where the bullet will be shot
Const BULLET_SHOT_TIME As Integer = 200 'Wait time between shots - in milliseconds
Const FRICTION As Single = 0.9905

'Variables
Dim shapes() As Shape 'Holds the shapes currently on the screen - dynamic array
Dim bullets() As Bullet 'Holds the bullets currently in play - dynamic array
Dim player As Ship 'The player's ship
Dim cameraX As Long 'The X position of the view on the screen
Dim cameraY As Long 'The Y position of the view on the screen
Dim scoreMultiplier As Integer 'How much to multiply the score by
Dim score As Long 'The player's score
Dim lifeScore As Integer 'The player's score on the current life
Dim boardSizeX As Long 'How wide the board is
Dim boardSizeY As Long 'How tall the board is
Dim level As Integer 'What level the player is on
Dim lives As Integer 'How many lives the player has
Dim bombs As Integer 'How many bombs the player has
Dim recordHighScore As Boolean 'Should the user's score should be checked for a high score?
Dim keyboard(1 To 255) As Boolean 'Represents ASCII codes of keyboard
Dim lastShot As Long 'What time the player last shot a bullet
Dim spawnInterval As Integer
Dim life As Long
Dim canBomb As Boolean

Private Sub Form_Activate()

    '1. Start the game after you can see the form
    Call startGame
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    '1. Set the value of the index of the keyboard code to True
    keyboard(KeyCode) = True
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    '1. Set the value of the index of the keyboard code to False
    keyboard(KeyCode) = False
    
    '2. Check to see if enter was just pressed (Don't want user to use multiple bombs)
    If KeyCode = vbKeyShift And bombs > 0 And canBomb Then
        'Use a bomb
        
        'Make the screen white temporarily
        Dim t As Long
        t = GetTickCount + 100
        frmGeometryWars.BackColor = vbWhite
        Do
            DoEvents
        Loop Until t < GetTickCount
        
        '2.1. Destroy all the shapes
        Do Until UBound(shapes) = 1
            Unload imgShape(shapes(2).controlIndex)
            removeShapeElement shapes, 2
        Loop
        
        'Change the background back to black
        frmGeometryWars.BackColor = vbBlack
        
        'Decrease the number of bombs
        bombs = bombs - 1
        lblBombs = bombs
    End If
    
End Sub

Private Sub Form_Load()
    'Start the game - set any variables to their appropriate values
    
    '1.  Seed the random number generator
    Randomize
    
    '2. Set the size of the Bullet and Shape arrays to (1 to 1)
    ReDim shapes(1 To 1)
    ReDim bullets(1 To 1)
    
    '3. Set the player.x and player.y so they match the position of the ship on the screen
    player.X = imgShip.left
    player.Y = imgShip.top
    
    'Start game is called in Form_Activate
    
End Sub

Sub startGame()
    'Reset variables for a new game
    
    Dim line As String 'To read from the file
    
    'Open the settings file to get the settings from
    Open "data\settings.txt" For Input As #1
    
    '1.  Set the score to 0
    score = 0
    lblScore = "Score: 0"
    
    'Set the level to 1
    level = 1
    
    '2. Set the score multiplier to 1
    scoreMultiplier = 1
    
    '2.  Set the recordHighScore flag to True
    recordHighScore = True
    
    '3.  Get the size of the screen from the settings
    Input #1, line
    boardSizeX = Int(line)
    Input #1, line
    boardSizeY = Int(line)
    
    '4.  Get the number lives the user wants to start off with from the settings file
    Input #1, line
    lives = Int(line)
    lblLives = ": " & lives
    
    '5.  If the numbers of lives is greater than 3 then
    If lives > 3 Then
        '5.1.     Set the recordHighScore flag to False
        recordHighScore = False
    End If
    
    '6.  Get the number of bombs the user wants to start off with
    Input #1, line
    bombs = Int(line)
    lblBombs = ": " & bombs
    
    '7.  If the number of bombs is greater than 3 then
    If bombs > 3 Then
        '7.1.    Set the recordHighScore flag to False
        recordHighScore = False
    End If
    lastShot = Now()
    
    '8.  Enable all the timers(tmrMoveShapes, tmrMoveShip, tmrSpawnShapes, tmrMoveBullets, tmrKeyboardInput)
    tmrMoveShip.Enabled = True
    tmrSpawnShapes.Enabled = True
    tmrMoveBullets.Enabled = True
    tmrKeyboardInput.Enabled = True
    tmrMoveShapes.Enabled = True
    
    Dim i As Integer
    For i = 1 To UBound(keyboard)
        keyboard(i) = False
    Next i
    
    player.xVel = 0
    player.yVel = 0
    player.rotation = 0
    
    '9. Change the camera's X and Y to 0
    cameraX = cameraY = 0
    
    'Change the outline co-ordinates
    Call setOutline
    
    'Set the spawn interval
    spawnInterval = 1200
    
    'Set the score needed for one life
    life = 5000
    
    canBomb = True
    
    'Close the file
    Close #1

End Sub

Private Sub tmrKeyboardInput_Timer()
    'Game controls
    
    '1. If the up arrow is pressed then
    If keyboard(vbKeyUp) Then
        '1.1. Decrease the Y velocity of the player
        player.yVel = player.yVel - PLAYER_SPEED
        
        '1.2. If the Y velocity is greater than the limit then
        If player.yVel < -MAX_PLAYER_SPEED Then
            '1.2.1. Change it back to what it was before
            player.yVel = player.yVel + PLAYER_SPEED
        End If
    End If
    
    '2. If the down arrow is pressed then
    If keyboard(vbKeyDown) Then
        '2.1. Increase the Y velocity of the player
        player.yVel = player.yVel + PLAYER_SPEED
        
        '2.2. If the Y velocity is greater than the limit then
        If player.yVel > MAX_PLAYER_SPEED Then
            '2.2.1. Change it back to what it was before
            player.yVel = player.yVel - PLAYER_SPEED
        End If
    End If
    
    
    
    '3. If the left arrow is pressed then
    If keyboard(vbKeyLeft) Then
        '3.1. Decrease the X velocity of the player
        player.xVel = player.xVel - PLAYER_SPEED
        
        '3.2. If the Y velocity is greater than the limit then
        If player.xVel < -MAX_PLAYER_SPEED Then
            '3.2.1. Change it back to what it was before
            player.xVel = player.xVel + PLAYER_SPEED
        End If
    End If
    
    '4. If the right arrow is pressed then
    If keyboard(vbKeyRight) Then
        '4.1. Increase the X velocity of the player
        player.xVel = player.xVel + PLAYER_SPEED
        
        '4.2. If the Y velocity is greater than the limit then
        If player.xVel > MAX_PLAYER_SPEED Then
            '4.2.1. Change it back to what it was before
            player.xVel = player.xVel - PLAYER_SPEED
        End If
    End If
    
    '5. If the 'D' key is pressed then
    If keyboard(vbKeyD) Then
    
        '5.1. Rotate the ship clockwise
        player.rotation = player.rotation + 2
        
        '5.2 If the ship is rotation more than 359 degrees then
        If player.rotation > 359 Then
        
            '5.2.1. Change the amount of rotation to 0
            player.rotation = 0
        End If
    End If
    
    '6.  If the 'A' key is pressed then
    If keyboard(vbKeyA) Then
    
        '6.1. Rotate the ship counter clockwise
        player.rotation = player.rotation - 2
        
        '6.2 If the ship is rotation less than 0 degrees then
        If player.rotation < 0 Then
        
            '6.2.1. Change the amount of rotation to 359
            player.rotation = 359
        End If
    End If
    
    '9.  If the spacebar is pressed then
    If keyboard(vbKeySpace) Then
    
        '9.1. Shoot a bullet in the direction that the ship is faced
        If lastShot + BULLET_SHOT_TIME < GetTickCount Then
            Call shootBullet
            lastShot = GetTickCount
        End If
    End If
    
End Sub

Private Sub tmrMoveBullets_Timer()
    'Update the positions of the bullets
    
    'Counters
    Dim i As Integer
    Dim j As Integer
    Dim collided As Boolean 'Variable to keep track of collision

    '1.  Cycle through all of the bullets in the bullet array
    i = 2
    Do Until i > UBound(bullets)

        '1.1.    Add the x velocity to the x co-ordinate
        bullets(i).X = bullets(i).X + bullets(i).xVel
        
        '1.2.    Add the y velocity to the y co-ordinate
        bullets(i).Y = bullets(i).Y + bullets(i).yVel
        
        'If the bullet travels off the screen then
        If bullets(i).X + imgBullet(1).Width > boardSizeX Or bullets(i).Y + imgBullet(1).Height > boardSizeY _
            Or bullets(i).X < 0 Or bullets(i).Y < 0 Then
        
            'Destroy the bullet
            Unload imgBullet(bullets(i).controlIndex)
            removeBulletElement bullets, i
        Else
            
            'Set the collided flag to false
            collided = False
            
            '1.3.    Cycle through all of the shapes in the shape array
            j = 2
            Do Until j > UBound(shapes)
                '1.3.1.  If the bullet hit the shape(bulletCollision) then
                If bulletCollision(bullets(i), shapes(j)) Then
                
                    'Update the score
                    score = score + (((shapes(j).type + 1) * 100) * scoreMultiplier)
                    
                    'Update the score on the current life
                    lifeScore = lifeScore + (((shapes(j).type + 1) * 100) * scoreMultiplier)
                    
                    'If the score on the current life is greater than what is needed for a new level then
                    If lifeScore >= life Then
                        'Make it harder!
                        
                        'Update the score multiplier
                        scoreMultiplier = scoreMultiplier + 1
                        
                        lblScoreMultiplier = "Score multiplier: " & scoreMultiplier
                        
                        'Update the level
                        level = level + 1
                        lblLevel = "Level: " & level
                        
                        'Change the life score back to 0
                        lifeScore = 0
                        
                        'Add on to what is needed for the next level
                        life = life + 3500
                        
                        'Update how fast the shapes spawn
                        spawnInterval = spawnInterval - 100
                        'If the spawn interval is less or equal to 0 then
                        If spawnInterval <= 0 Then
                            'Change it to 1 - the fastest
                            spawnInterval = 1
                        End If
                        Label1 = spawnInterval
                        tmrSpawnShapes.Interval = spawnInterval
                        
                    End If
                    
                    'Update the score label
                    lblScore = "Score: " & score
                    
                    '1.3.1.1.     Destroy the bullet
                    Unload imgBullet(bullets(i).controlIndex)
                    removeBulletElement bullets, i
                    
                    'If the shape is a spinner, make a tiny spinner
                    If shapes(j).type = SpinnerShape Then
                        shapes(j).type = TinySpinnershape
                        imgShape(shapes(j).controlIndex) = imgShapeImages(TinySpinnershape)
                        
                    'Otherwise
                    Else
                    
                        '1.3.1.2.     Destroy the shape
                        Unload imgShape(shapes(j).controlIndex)
                        removeShapeElement shapes, j
                    
                        'Set the collided flag to true
                        collided = True
                        
                    End If
                    
                    'Stop checking for collision
                    Exit Do
                Else
                
                    'Only update the counter if there wasn't a collision because the array is resized
                    j = j + 1
                    
                End If
                
            Loop
            
            'If there was no collision then
            If Not collided Then
                'Update the counter
                i = i + 1
            End If
        End If
        
    Loop

End Sub

Private Sub tmrMoveShapes_Timer()
    'Update the positions of the shapes, depending on their AI
    
    Dim i As Integer
    
    '1.  Cycle through all of the shapes in the shape array
    For i = 2 To UBound(shapes)
    
        '1.1.    If the ShapeType is a Wanderer then
        If shapes(i).type = WandererShape Then
            '1.1.1.  Call the WandererAI
            Call WandererAI(shapes(i))
        'Repeat 1.1 for all shapes
        ElseIf shapes(i).type = GruntShape Then
            Call GruntAI(shapes(i))
        ElseIf shapes(i).type = WeaverShape Then
            Call WeaverAI(shapes(i))
        ElseIf shapes(i).type = SpinnerShape Then
            Call SpinnerAI(shapes(i))
        ElseIf shapes(i).type = TinySpinnershape Then
            Call SpinnerAI(shapes(i))
        End If
        
        'Update the co-ordinates of the shape
        shapes(i).X = shapes(i).X + shapes(i).xVel
        shapes(i).Y = shapes(i).Y + shapes(i).yVel
        
        '1.2.    If the shape has collided with the ship(shipCollision) then
        If shipCollision(shapes(i)) Then
        
            '1.2.1.   Call lostLife
            Call lostLife
            
            Exit For
        End If
        
    Next i
    'tmrMoveBullets handles shape and bullet collision

End Sub

Private Sub tmrMoveShip_Timer()
    'Move the ship based on it's x and y velocities
    'tmrMoveShapes handles ship and shape collision
    
    '1.  Add the x velocity to the x co-ordinate
    player.X = player.X + player.xVel
    
    '2.  If the x co-ordinate is beyond the boundaries of the screen(less than 0 or greater than the width) then
    If player.X < 0 Or player.X + imgShip.Width > boardSizeX Then
        '2.1.     Make the x velocity negative of what it was before
        player.xVel = -player.xVel
    End If
    
    '3.  Add the y velocity to the y co-ordinate
    player.Y = player.Y + player.yVel
    
    '4.  If the y co-ordinate is beyond the boundaries of the screen(less than 0 or greater than the height) then
    If player.Y < 0 Or player.Y + imgShip.Height > boardSizeY Then
        '4.1.     Make the y velocity negative of what it was before
        player.yVel = -player.yVel
    End If

    'To stop the player from moving gradually
    player.xVel = player.xVel * FRICTION
    player.yVel = player.yVel * FRICTION
    
    '5.  Centre the camera over the ship
    cameraX = Int(player.X - (frmGeometryWars.Width / 2))
    cameraY = Int(player.Y - (frmGeometryWars.Height / 2))

    'Adjust the position of the ship based on the camera
    'Adjust the position of the boundaries based on the camera
    Call changePositions

End Sub

Sub changePositions()
    'Update the co-ordinates of the controls
    
    Dim i As Integer 'Counter
    
    'Find the new ship co-ordinates
    imgShip.left = player.X - cameraX
    imgShip.top = player.Y - cameraY
    
    'Figure out the position of the shooting line
    Dim angle As Integer
    If (player.rotation >= 270 And player.rotation <= 360) _
        Or (player.rotation >= 90 And player.rotation <= 180) Then
        angle = (player.rotation Mod 90)
    Else
        angle = 90 - (player.rotation Mod 90)
    End If
    
    Dim xLength As Integer
    xLength = Cos(angle / 57.29578) * LINE_LENGTH
    Dim ylength As Integer
    ylength = -(Sin(angle / 57.29578) * LINE_LENGTH)
    If player.rotation > 270 Then
        xLength = -xLength
    ElseIf player.rotation > 180 Then
        xLength = -xLength
        ylength = -ylength
    ElseIf player.rotation > 90 Then
        ylength = -ylength
    End If
    linShoot.X1 = imgShip.left + (imgShip.Width / 2)
    linShoot.Y1 = imgShip.top + (imgShip.Height / 2)
    linShoot.X2 = linShoot.X1 + xLength
    linShoot.Y2 = linShoot.Y1 + ylength
    
    
    'Find all of the shape co-ordinates
    For i = 2 To UBound(shapes)
        imgShape(shapes(i).controlIndex).left = shapes(i).X - cameraX
        imgShape(shapes(i).controlIndex).top = shapes(i).Y - cameraY
    Next i
    
    'Find all of the bullet co-ordinates
    For i = 2 To UBound(bullets)
        imgBullet(bullets(i).controlIndex).left = bullets(i).X - cameraX
        imgBullet(bullets(i).controlIndex).top = bullets(i).Y - cameraY
    Next i
    
    'Set the outline
    Call setOutline
End Sub

Sub setOutline()
    'Uses camerax and cameraY to change the co-ordinates of the outline
    'Note: Outline is used to prevent flickering inside the arena
    
    'top -
    linOutline(1).Y1 = -cameraY
    linOutline(1).X1 = -cameraX
    linOutline(1).Y2 = -cameraY
    linOutline(1).X2 = -cameraX + boardSizeX
    
    'left |
    linOutline(2).Y1 = -cameraY
    linOutline(2).X1 = -cameraX
    linOutline(2).Y2 = -cameraY + boardSizeY
    linOutline(2).X2 = -cameraX
    
    'right |
    linOutline(3).Y1 = -cameraY
    linOutline(3).X1 = -cameraX + boardSizeX
    linOutline(3).Y2 = -cameraY + boardSizeY
    linOutline(3).X2 = -cameraX + boardSizeX
    
    'bottom -
    linOutline(4).Y1 = -cameraY + boardSizeY
    linOutline(4).X1 = -cameraX
    linOutline(4).Y2 = -cameraY + boardSizeY
    linOutline(4).X2 = -cameraX + boardSizeX
    
End Sub

Private Sub tmrSpawnShapes_Timer()
    'Spawn new shapes at random timer intervals
    
    Dim shapeNumber As Integer 'The shape to create
    Dim arrayIndex As Integer 'The array index to use for the shape array
    Dim controlIndex As Integer
    
    '1.  If the level is above the number of shapes then
    If level > TinySpinnershape Then
        '1.1.     Create a random number from 1 to the number of shapes
        Do
            shapeNumber = Int(Rnd * (TinySpinnershape + 1))
        Loop Until shapeNumber <> TinySpinnershape
        
    '2.  Otherwise if the level is less than or equal to the number of shapes then
    ElseIf level <= TinySpinnershape Then
        '2.1.     Create a random number from 1 to the current level
        Do
            shapeNumber = Int(Rnd * level)
        Loop Until shapeNumber <> TinySpinnershape
    End If
    
    
    'Find the array index
    arrayIndex = UBound(shapes) + 1
    'Find the control index
    controlIndex = imgShape.UBound + 1
    
    '3.  Add a new shape to the shapes array and the control array
    ReDim Preserve shapes(1 To arrayIndex)
    Load imgShape(controlIndex)
    
    'Set the visible property to true
    imgShape(controlIndex).Visible = True
    
    'Depending the random number chosen, change the picture
    imgShape(controlIndex) = imgShapeImages(shapeNumber)
    
    '4.  Give it the type of the random number
    shapes(arrayIndex).type = shapeNumber
    
    'Set the control index because the arrayIndex and controlIndex don't match
    shapes(arrayIndex).controlIndex = controlIndex
    
    '5.  Assign a random x and y within the screen until it is far enough away from the ship
    Do
        shapes(arrayIndex).X = Int(Rnd * (boardSizeX - imgShape(controlIndex).Width))
        shapes(arrayIndex).Y = Int(Rnd * (boardSizeY - imgShape(controlIndex).Height))
        imgShape(controlIndex).left = shapes(arrayIndex).X - cameraX
        imgShape(controlIndex).top = shapes(arrayIndex).Y - cameraY
    Loop Until shapes(arrayIndex).X + imgShape(controlIndex).Width < player.X - 500 _
        Or shapes(arrayIndex).X > player.X + imgShip.Width + 500 Or _
        shapes(arrayIndex).Y + imgShape(controlIndex).Height < player.Y - 500 Or _
        shapes(arrayIndex).Y > player.Y + imgShip.Height + 500

End Sub

Sub shootBullet()
    'What happens when the player shoots a bullet
    
    Dim arrayIndex As Integer
    '1.  Resize the 'bullet' array and add a new bullet
    arrayIndex = UBound(bullets) + 1
    ReDim Preserve bullets(1 To arrayIndex)
    
    '2.  Set the x and y values of the bullet to the x and y values of the ship
    bullets(arrayIndex).X = player.X + (imgShip.Width / 2)
    bullets(arrayIndex).Y = player.Y + (imgShip.Height / 2)
    
    '3.  Load a new image bullet control into the bullet control array
    Load imgBullet(imgBullet.UBound + 1)
    
    '4.  Set the top and left co-ordinates to the same x and y values of the ship
    imgBullet(imgBullet.UBound).top = imgShip.top + (imgShip.Height / 4)
    imgBullet(imgBullet.UBound).left = imgShip.left + (imgShip.Width / 4)
    imgBullet(imgBullet.UBound).Visible = True
    bullets(arrayIndex).controlIndex = imgBullet.UBound
    
    '5.  Figure out the x velocity and y velocity depending on the angle the ship is faced at
    '5.1.    Use trigonometric ratios to find slope and move along the line
    Dim angle As Integer
    If (player.rotation > 270 And player.rotation < 360) _
        Or (player.rotation > 90 And player.rotation < 180) Then
        angle = (player.rotation Mod 90)
    Else
        angle = 90 - (player.rotation Mod 90)
    End If
    
    Dim xLength As Integer
    xLength = Cos(angle / 57.29578) * LINE_LENGTH
    Dim ylength As Integer
    ylength = -(Sin(angle / 57.29578) * LINE_LENGTH)
    
    If player.rotation > 270 Then
        xLength = -xLength
    ElseIf player.rotation > 180 Then
        xLength = -xLength
        ylength = -ylength
    ElseIf player.rotation > 90 Then
        ylength = -ylength
    End If
    
    bullets(arrayIndex).xVel = xLength / 15
    bullets(arrayIndex).yVel = ylength / 15
    
End Sub

Private Sub removeBulletElement(a() As Bullet, arrayIndex As Integer)
    'Resize a dynamic array by removing the specified element
    
    Dim i As Integer
    '1.  Cycle through the elements from the arrayIndex to the second last element of the array
    For i = arrayIndex To UBound(a) - 1
        '1.1.    Change the value at the current position to the value of the position one element ahead in the array
        a(i) = a(i + 1)
    Next i
    
    '2.  Resize the array to what it's size was before, subtract one element
    ReDim Preserve a(1 To UBound(a) - 1)
    
End Sub
Private Sub removeShapeElement(a() As Shape, arrayIndex As Integer)
    'Resize a dynamic array by removing the specified element
    
    Dim i As Integer
    '1.  Cycle through the elements from the arrayIndex to the second last element of the array
    For i = arrayIndex To UBound(a) - 1
        '1.1.    Change the value at the current position to the value of the position one element ahead in the array
        a(i) = a(i + 1)
    Next i
    
    '2.  Resize the array to what it's size was before, subtract one element
    ReDim Preserve a(1 To UBound(a) - 1)
    
End Sub

Private Function bulletCollision(b As Bullet, s As Shape) As Boolean
    'Collision between a bullet and a shape
    
    Dim bulletWidth As Integer
    Dim bulletHeight As Integer 'The dimensions of the bullet
    Dim shapeWidth As Integer
    Dim shapeHeight As Integer 'The dimensions of the shape
    
    '1.  Set the bulletWidth to the width of the bullet image control
    bulletWidth = imgBullet(1).Width
    
    '2.  Set the bulletHeight to the height of the bullet image control
    bulletHeight = imgBullet(1).Height
    
    '3.  If the ShapeType is a Wanderer then
    '3.1.    Set the shapeWidth to the width of the Wanderer image control
    shapeWidth = imgShapeImages(s.type).Width
        
    '3.2.    Set the shapeHeight to the height of the Wanderer image control
    shapeHeight = imgShapeImages(s.type).Height
    
    '4.  If the x co-ordinate of the shape is less than the x co-ordinate of the bullet + the
    'width of the bullet and the x co-ordinate of the shape + the width of the shape is greater
    'than the x co-ordinate of the bullet and the y co-ordinate of the shape is less than the
    'y co-ordinate of the bullet + the height of the bullet and the y co-ordinate of the shape
    '+ the height of the shape is greater than the y co-ordinate of the bullet then
    If s.X < b.X + bulletWidth And s.X + shapeWidth > b.X And s.Y < b.Y + bulletHeight _
        And s.Y + shapeHeight > b.Y Then
        
        '4.1.    There was a collision
        bulletCollision = True
    End If

End Function

Private Function shipCollision(s As Shape) As Boolean
    'Collision between a ship and a shape
    
    Dim shapeWidth As Integer
    Dim shapeHeight As Integer 'The dimensions of the shape
    
    'Set the shape width and shape height depending on the type
    shapeWidth = imgShapeImages(s.type).Width
    shapeHeight = imgShapeImages(s.type).Height
    
    '2.  If the x co-ordinate of the shape is less than the x co-ordinate
    'of the ship + the width of the ship and the x co-ordinate of the shape
    '+ the width of the shape is greater than the x co-ordinate of the ship
    'and the y co-ordinate of the shape is less than the y co-ordinate of the ship
    '+ the height of the ship and the y co-ordinate of the shape + the height of the
    'shape is greater than the y co-ordinate of the ship then
    If s.X < player.X + imgShip.Width And s.X + shapeWidth > player.X _
        And s.Y < player.Y + imgShip.Height And s.Y + shapeHeight > player.Y Then
        '2.1.    There was a collision
        shipCollision = True
    End If
    
    'No collision was detected
    
End Function

Sub lostLife()
    'What happens when you lose a life
    
    'Counter
    Dim i As Integer
    
    '1.  Unload all elements (except 1) in the shape control array
    For i = 2 To UBound(shapes)
        Unload imgShape(shapes(i).controlIndex)
    Next i
        
    '2.  Unload all elements (except 1) in the bullet control array
    For i = 2 To UBound(bullets)
        Unload imgBullet(bullets(i).controlIndex)
    Next i
    
    '3.  Change the size of the shape array to (1 to 1)
    ReDim shapes(1 To 1)
    
    '4.  Change the size of the bullet array to (1 to 1)
    ReDim bullets(1 To 1)
    
    'Take a life away
    lives = lives - 1
    
    'If the player has no live left then
    If lives = 0 Then
        'They lose
        MsgBox "You lose!", vbOKOnly, "Geometry Wars"
        
        'The game is over
        Call gameOver
    'Otherwise
    Else
        canBomb = False
        
        'They lost a life
        MsgBox "Lost a life!", vbOKOnly, "Geometry Wars"
        'Reset the co-ordinates of the ship
        player.X = 1000
        player.Y = 1000
        'Reset the x and y velocities of the ship
        player.xVel = 0
        player.yVel = 0
        'Reset the rotation
        player.rotation = 0
        'Reset the score multiplier
        scoreMultiplier = 1
        lblScoreMultiplier = "Score multiplier: 1"
        'Change the amount of score on the current life to 0
        lifeScore = 0
        'Change what is needed to get to another level
        life = 5000 + (level * 200)
        'Change the spawn interval
        spawnInterval = 1200 - (level * 20)
        tmrSpawnShapes.Interval = spawnInterval
        Label1 = spawnInterval
        For i = 1 To UBound(keyboard)
            keyboard(i) = False
        Next i
        lblLives = ": " & lives
        canBomb = True
    End If
    
    
End Sub

Sub gameOver()
    'What to do when the user lost
    
    Dim i As Integer
    
    '1.  Disable all the timers(tmrMoveShapes, tmrMoveShip, tmrSpawnShapes, tmrMoveBullets, tmrKeyboardInput)
    tmrMoveShapes.Enabled = False
    tmrMoveShip.Enabled = False
    tmrSpawnShapes.Enabled = False
    tmrMoveBullets.Enabled = False
    tmrKeyboardInput.Enabled = False
    
    '2.  If the recordHighScore flag is True
    If recordHighScore Then
        '2.1.    Call checkHighScored
        Call checkHighScore
    End If
    
    'Go back to the menu
    Unload Me
    frmMenu.Show
    
End Sub

Sub checkHighScore()
    'Check to see if the player got a high score
    '5 high scores will be held in the file
    
    Dim scores(1 To 5) As Long 'The scores in the high score file
    Dim names(1 To 5) As String 'The names in the high score file
    Dim newScores(1 To 5) As Long 'What scores will go in the high score file
    Dim newNames(1 To 5) As String 'What names will go in the high score file
    Dim position As Integer 'Where the player's new high score will go
    Dim i As Integer
    Dim line As String
    
    '1.  Open the high score file for input
    Open "data\highscores.txt" For Input As #1
    '2.  Cycle through all the lines in the high score file
    For i = 1 To 5
        Input #1, line
        
        '2.1.    Add the name into the names array
        names(i) = line
        
        Input #1, line
        
        '2.2.    Add the score into the scores array
        scores(i) = Int(line)
        
        '2.3.    Continue until the EOF is reached
        'Should always be 10 lines
    Next i
    
    '3.  Close the high score file
    Close #1
    
    '4.  Set the position to 0
    position = 0
    
    '5.  Cycle through all of the high scores
    For i = 1 To 5
    
        '5.1.    If the player's score is higher than the current high score then
        If score > scores(i) Then
        
            '5.1.1.   Set the position to the index in the array
            position = i
            
            '5.1.2.   Exit out of the for loop
            Exit For
        End If
    Next i
    
    '6.  If the position isn't 0 then
    If position <> 0 Then
        '6.1.    Get the player's name
        Dim name As String
        name = InputBox("New high score! Enter your name:", "Geometry Wars")
        If name <> "" Then
            For i = 1 To position - 1
                '6.2.1.   Change the value at the current position in the newScores array to what the
                'value was in the scores array in the element before
                newScores(i) = scores(i)
                newNames(i) = names(i)
            Next i
            
            '6.2.    Cycle through the scores from the position + 1 to the end of the scores array
            For i = position + 1 To 5
                '6.2.1.   Change the value at the current position in the newScores array to what the
                'value was in the scores array in the element before
                newScores(i) = scores(i - 1)
                newNames(i) = names(i - 1)
            Next i
            '6.3.    Change the score at 'position' in the newScores array to what the player's score is
            newScores(position) = score
            
            '6.4.    Change the name at 'position' in the newNames array to what the player entered as their name
            newNames(position) = name
            
            '6.5.    Open the high score file for output
            Open "data\highscores.txt" For Output As #1
            
            '6.6.    Cycle through the names and scores in newNames and newScores
            For i = 1 To 5
                '6.6.1.   Print the current name into the high score file
                Print #1, newNames(i)
                
                '6.6.2.   Print the current score into the high score file
                Print #1, newScores(i)
            Next i
            '6.7.    Close the high score file
            Close #1
        End If
    End If
End Sub

Private Sub WandererAI(s As Shape)
    'Wanders randomly around the board
    
    If s.xVel = 0 Or s.yVel = 0 Then
        Dim randNum As Integer
        '1.  Get a random number from 1 to 2
        randNum = Int(Rnd * 2) + 1
        
        '2.  If the number is 1 then
        If randNum = 1 Then
            '2.1.     Change the x velocity of the shape to 1
            s.xVel = 10
        '3.  Otherwise
        Else
            '3.1.     Change the x velocity of the shape to  -1
            s.xVel = -10
        End If
    'Repeat steps 1-3 for the y speed
        randNum = Int(Rnd * 2) + 1
        
        '2.  If the number is 1 then
        If randNum = 1 Then
            '2.1.     Change the y velocity of the shape to 1
            s.yVel = 10
        '3.  Otherwise
        Else
            '3.1.     Change the y velocity of the shape to  -1
            s.yVel = -10
        End If
    Else
        If s.Y < 0 Or s.Y + imgShape(s.controlIndex).Height > boardSizeY Then
            s.yVel = -s.yVel
        End If
        If s.X < 0 Or s.X + imgShape(s.controlIndex).Width > boardSizeX Then
            s.xVel = -s.xVel
        End If
    End If
End Sub

Private Sub GruntAI(s As Shape)
    'Slowly follows the ship
    
    Dim diffX As Long
    Dim diffY As Long 'Differences of ship to shape
    Dim distance As Single 'Distance from ship to shape
    Dim speed As Integer 'Speed of the shape
    
    'WeaverAI uses this, so we need to check what type it is
    If s.type = WeaverShape Then
        speed = 25
    Else
        speed = 15
    End If
    
    '1.  Find out the distance the shape is from the ship
    diffX = player.X - s.X
    diffY = player.Y - s.Y
    
    '1.1.    Pythagoras Theorem
    distance = Sqr(diffX ^ 2 + diffY ^ 2)
    
    '2.  Move along the shortest line from the ship to the shape
    s.xVel = (diffX * speed) / distance
    s.yVel = (diffY * speed) / distance

End Sub

Private Sub SpinnerAI(s As Shape)
    'Follows the ship at a high speed - breaks into tiny spinners when destroyed
    'Need to check to see if it moves out of the way the same way as Weavers do,
    'or if they just move the same as Grunts, just faster
    
    Dim diffX As Long
    Dim diffY As Long
    Dim distance As Single
    
    '1.  Find out the distance the shape is from the ship
    diffX = player.X - s.X
    diffY = player.Y - s.Y
    '1.1.    Pythagoras Theorem
    distance = Sqr(diffX ^ 2 + diffY ^ 2)
        
    '2.  Move along the shortest line from the ship to the shape, faster than a grunt
    'Formula: (deltaX * speed) / distance normalizes speed no matter what the slope or distance is
    s.xVel = (diffX * 30) / distance
    s.yVel = (diffY * 30) / distance
End Sub

Private Sub WeaverAI(s As Shape)
    'Follows the ship and avoids bullets being shot at it
    
    Dim i As Integer
    Dim j As Integer
    Dim b As Bullet
    Dim avoid As Boolean
    
    '1.  Cycle through all the bullets
    For i = 2 To UBound(bullets)
        
        'Set the temporary bullets X and Y values
        b.X = bullets(i).X
        b.Y = bullets(i).Y
        avoid = False
        
        'Loop a random number of times
        For j = 1 To 35
            'Pretending the bullet is moving
            b.X = b.X + (bullets(i).xVel * 2)
            b.Y = b.Y + (bullets(i).yVel * 2)
            
            'If on the next bullets movement it will collide with the shape then
            If bulletCollision(b, s) Then
                'Move the shape away from the ship
                If b.xVel < 0 Then
                    s.xVel = 55
                Else
                    s.xVel = -55
                End If
                
                If b.yVel < 0 Then
                    s.yVel = 55
                Else
                    s.xVel = -55
                End If
                
                avoid = True
                
                Exit For
            End If
        Next j
    Next i
    
    '2. If there is no bullet moving towards the shape then
    If Not avoid Then
        'Move towards the ship
        Call GruntAI(s)
    End If
    
End Sub

