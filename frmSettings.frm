VERSION 5.00
Begin VB.Form frmSettings 
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
   Begin VB.CommandButton cmdDefaults 
      BackColor       =   &H0000FF00&
      Caption         =   "Change to defaults"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtBombs 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      MaxLength       =   1
      TabIndex        =   8
      Top             =   6840
      Width           =   2295
   End
   Begin VB.TextBox txtLives 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      MaxLength       =   1
      TabIndex        =   7
      Top             =   5760
      Width           =   2295
   End
   Begin VB.TextBox txtScreenHeight 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   5040
      Width           =   2295
   End
   Begin VB.TextBox txtScreenWidth 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   5
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdMenuReturn 
      BackColor       =   &H0000FF00&
      Caption         =   "Return to menu"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   600
      Width           =   1935
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
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   1695
      Left            =   4920
      TabIndex        =   10
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Image imgTitle 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   3480
      Picture         =   "frmSettings.frx":0000
      Top             =   600
      Width           =   7860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number of bombs:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5160
      TabIndex        =   4
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label lblLives 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number of lives:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label lblScreenHeight 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Screen height:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label lblScreenWidth 
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Screen width:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   4320
      Width           =   1935
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDefaults_Click()

    'Set all of the text box values back to their defaults
    txtScreenWidth = 19200
    txtScreenHeight = 12000
    txtLives = 3
    txtBombs = 3
    
End Sub

Private Sub cmdMenuReturn_Click()
    
    Dim decision As Integer 'If they want to save their settings
    decision = vbYes 'Default to yes because you want to continue if the settings are default
    
    'If one of the settings does not match the default then
    If txtScreenWidth <> 19200 Or txtScreenHeight <> 12000 Or txtBombs > 3 Or txtLives > 3 Then
        'Tell them their high score won't be saved
        decision = MsgBox("Warning: With these settings, your high score will not be saved. Do you wish to continue?", vbYesNo, "Geometry Wars")
    End If
    
    'If they want to continue then
    If decision = vbYes Then
        '1. Open the settings file to save the settings
        Open "data\settings.txt" For Output As #1
        
        '2. Save the new settings
        Print #1, txtScreenWidth
        Print #1, txtScreenHeight
        Print #1, txtLives
        Print #1, txtBombs
        
        '3. Close the file
        Close #1
        
        '3. Return to the menu
        Unload frmSettings
        frmMenu.Show
    End If
    
End Sub

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

Private Sub Form_Load()

    Dim line As String 'For reading from the file
    'Format for settings file:
    'Line 1: Screen width
    'Line 2: Screen height
    'Line 3: Number of lives
    'Line 4: Number of bombs
    
    '1. Load settings from the settings file
    Open "data\settings.txt" For Input As #1
    
    '2. Set the text of all of the text boxes to what it is in the settings file
    'Screen width
    Input #1, line
    txtScreenWidth = line
    
    'Screen height
    Input #1, line
    txtScreenHeight = line
    
    'Number of lives
    Input #1, line
    txtLives = line
    
    'Number of bombs
    Input #1, line
    txtBombs = line
    
    '3. Close the file
    Close #1
    
End Sub

