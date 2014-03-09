VERSION 5.00
Begin VB.Form frmHighScores 
   BackColor       =   &H80000007&
   Caption         =   "Geometry Wars"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   1935
   End
   Begin VB.Line linLines 
      BorderColor     =   &H8000000D&
      Index           =   1
      X1              =   360
      X2              =   360
      Y1              =   0
      Y2              =   11160
   End
   Begin VB.Label lblLabelScore 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   8280
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label lblLabelName 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   5160
      TabIndex        =   11
      Top             =   3840
      Width           =   3015
   End
   Begin VB.Label lblHighScores 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "High Scores"
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
      Left            =   3840
      TabIndex        =   10
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   8280
      TabIndex        =   9
      Top             =   9360
      Width           =   1575
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   8280
      TabIndex        =   8
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   8280
      TabIndex        =   7
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   8280
      TabIndex        =   6
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   8280
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   5
      Left            =   5160
      TabIndex        =   4
      Top             =   9360
      Width           =   3015
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   4
      Left            =   5160
      TabIndex        =   3
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   3
      Left            =   5160
      TabIndex        =   2
      Top             =   7200
      Width           =   3015
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   2
      Left            =   5160
      TabIndex        =   1
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Index           =   1
      Left            =   5160
      TabIndex        =   0
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Image imgTitle 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   3480
      Picture         =   "frmHighScores.frx":0000
      Top             =   600
      Width           =   7860
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMenuReturn_Click()

    '1. Unload the current form
    Unload Me
    
    '2. Show the menu
    frmMenu.Show
    
End Sub

Private Sub Form_Activate()
    
    Dim i As Integer 'Counter
    Dim name As String, score As String
    
    '1. Open the high score file
    Open "data\highscores.txt" For Input As #1
    
    '2. Cycle through all the scores
    For i = 1 To 5
    
        '2.1. Get the name
        Input #1, name
        lblName(i) = name
        
        '2.2. Get the score
        Input #1, score
        lblScore(i) = score
        
    Next i
    
    '3. Close the high score file
    Close #1

    'In Form_Activate so that the screen Width and Height are set
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

