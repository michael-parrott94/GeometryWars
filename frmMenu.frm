VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H80000007&
   Caption         =   "Geometry Wars"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Shape shpHighScores 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   4800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblHighScores 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Highscores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   4200
      TabIndex        =   5
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1095
      Left            =   6360
      TabIndex        =   4
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Shape shpQuitDot 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   8040
      Shape           =   3  'Circle
      Top             =   8400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape shpSettingsDot 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6480
      Shape           =   3  'Circle
      Top             =   7200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape shpInstructionsDot 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape shpPlayDot 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      FillColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line linLines 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   360
      X2              =   360
      Y1              =   0
      Y2              =   11040
   End
   Begin VB.Label lblSettings 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   8880
      TabIndex        =   2
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   5760
      TabIndex        =   1
      Top             =   5880
      Width           =   3015
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Image imgTitle 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   3480
      Picture         =   "frmMenu.frx":0000
      Top             =   600
      Width           =   7860
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    'In Form_Activate so that the screen Width and Height are set
    
    Dim i As Integer 'Counter
    Dim left As Integer 'The X co-ordinates of the vertical lines
    left = linLines(1).X1
    Dim top As Integer 'The Y co-ordinates of the horizontal lines
    top = linLines(1).X1
    
    'Start off at the second index
    i = 2

    Do
        'Load the current line
        Load linLines(i)
        'Change the X co-ordinates to the current x co-ordinate
        linLines(i).X1 = left
        linLines(i).X2 = left
        
        'Have the line from the top of the screen to the bottom
        linLines(i).Y1 = 0
        linLines(i).Y2 = Me.Height
        
        'Make the line visible
        linLines(i).Visible = True
        
        'Update the left position
        left = left + linLines(1).X1
        
        'Update the counter
        i = i + 1
    'Continue until we are past the width of the screen
    Loop Until left > Me.Width
    
    Do
        'Load the current line
        Load linLines(i)
        
        'Set the Y co-ordinates of the line to the current y co-ordinate
        linLines(i).Y1 = top
        linLines(i).Y2 = top
        
        'Have the line go from the far left to the far right of the screen
        linLines(i).X1 = 0
        linLines(i).X2 = Me.Width

        'Make the line visible
        linLines(i).Visible = True
        
        'Update the y co-ordinate
        top = top + linLines(1).X1
        
        'Update the counter
        i = i + 1
        
    'Continue until you past the height of the screen
    Loop Until top > Me.Height
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'When the mouse is not on a label, show that nothing is selected
    
    lblQuit.ForeColor = vbWhite
    lblPlay.ForeColor = vbWhite
    lblHighScores.ForeColor = vbWhite
    lblInstructions.ForeColor = vbWhite
    lblSettings.ForeColor = vbWhite
    
    shpQuitDot.Visible = False
    shpPlayDot.Visible = False
    shpHighScores.Visible = False
    shpInstructionsDot.Visible = False
    shpSettingsDot.Visible = False
    
End Sub

Private Sub lblHighScores_Click()
    
    '1. Unload the current form
    Unload Me
    
    '2. Show the high scores form
    frmHighScores.Show
    
End Sub

Private Sub lblInstructions_Click()

    'Using Windows help, show the Geometry Wars help file
    Shell "HH.exe data/geometrywars.chm", vbMaximizedFocus
    
End Sub

Private Sub lblPlay_Click()
    
    '1. Unload the menu form
    Unload frmMenu
    
    '2. Show the Geometry Wars form - the game form
    frmGeometryWars.Show
    
End Sub

Private Sub lblQuit_Click()
    
    '1. Unload all of the forms
    Unload frmMenu
    Unload frmGeometryWars
    Unload frmSettings
    
    '2. End the program
    End
    
End Sub

Private Sub lblSettings_Click()

    '1. Hide the menu
    Unload Me
    frmMenu.Hide
    '2. Show the settings
    frmSettings.Show
    
End Sub

'For all of the following subprograms:
'Depending on where the mouse is, show that it is selected by:
'1. Making the text red
'2. Showing a dot beside the label
Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblPlay.ForeColor = vbRed
    lblHighScores.ForeColor = vbWhite
    lblInstructions.ForeColor = vbWhite
    lblSettings.ForeColor = vbWhite
    lblQuit.ForeColor = vbWhite
    
    shpPlayDot.Visible = True
    shpHighScores.Visible = False
    shpInstructionsDot.Visible = False
    shpSettingsDot.Visible = False
    shpQuitDot.Visible = False
    
End Sub

Private Sub lblHighScores_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblHighScores.ForeColor = vbRed
    lblPlay.ForeColor = vbWhite
    lblInstructions.ForeColor = vbWhite
    lblSettings.ForeColor = vbWhite
    lblQuit.ForeColor = vbWhite
    
    shpHighScores.Visible = True
    shpPlayDot.Visible = False
    shpInstructionsDot.Visible = False
    shpSettingsDot.Visible = False
    shpQuitDot.Visible = False
    
End Sub

Private Sub lblInstructions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblInstructions.ForeColor = vbRed
    lblPlay.ForeColor = vbWhite
    lblHighScores.ForeColor = vbWhite
    lblSettings.ForeColor = vbWhite
    lblQuit.ForeColor = vbWhite
    
    shpInstructionsDot.Visible = True
    shpPlayDot.Visible = False
    shpHighScores.Visible = False
    shpSettingsDot.Visible = False
    shpQuitDot.Visible = False
    
End Sub

Private Sub lblSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblSettings.ForeColor = vbRed
    lblPlay.ForeColor = vbWhite
    lblHighScores.ForeColor = vbWhite
    lblInstructions.ForeColor = vbWhite
    lblQuit.ForeColor = vbWhite
    
    shpSettingsDot.Visible = True
    shpPlayDot.Visible = False
    shpHighScores.Visible = False
    shpInstructionsDot.Visible = False
    shpQuitDot.Visible = False
    
End Sub

Private Sub lblQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblQuit.ForeColor = vbRed
    lblPlay.ForeColor = vbWhite
    lblHighScores.ForeColor = vbWhite
    lblInstructions.ForeColor = vbWhite
    lblSettings.ForeColor = vbWhite
    
    shpQuitDot.Visible = True
    shpPlayDot.Visible = False
    shpHighScores.Visible = False
    shpInstructionsDot.Visible = False
    shpSettingsDot.Visible = False
    
End Sub
