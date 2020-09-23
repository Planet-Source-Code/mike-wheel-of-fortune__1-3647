VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wheel-of-Fortune"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   Icon            =   "whhel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Solve"
      Height          =   375
      Left            =   8520
      TabIndex        =   45
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Spin"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   41
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   6720
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS LineDraw"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5880
      TabIndex        =   49
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS LineDraw"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   48
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   47
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00008000&
      X1              =   0
      X2              =   11880
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   615
      Left            =   1920
      Top             =   120
      Width           =   8295
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H0000C000&
      Height          =   375
      Left            =   8520
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000C000&
      Height          =   855
      Left            =   4320
      Top             =   5760
      Width           =   3975
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H0000C000&
      Height          =   255
      Left            =   1080
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000C000&
      Height          =   855
      Left            =   1080
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS LineDraw"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   46
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      Height          =   495
      Index           =   5
      Left            =   3360
      TabIndex        =   44
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label ass 
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5880
      TabIndex        =   43
      Top             =   1440
      Width           =   1725
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   8520
      TabIndex        =   42
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Wheel                       Of               Fortune"
      BeginProperty Font 
         Name            =   "StopD"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   1920
      TabIndex        =   40
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label lbl1 
      BackColor       =   &H00000000&
      Caption         =   "Letters Chosen-----------"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   1080
      TabIndex        =   39
      Top             =   5760
      Width           =   3015
   End
   Begin VB.Label lblLetters 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   1080
      TabIndex        =   38
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label lblMsg1 
      BackColor       =   &H00000000&
      Caption         =   "Pick a letter or type the answer         to sove the puzzle."
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   4320
      TabIndex        =   37
      Top             =   5760
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000012&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   0
      Top             =   5160
      Width           =   12255
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   8160
      TabIndex        =   36
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   7680
      TabIndex        =   35
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   7200
      TabIndex        =   34
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   6720
      TabIndex        =   33
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6240
      TabIndex        =   32
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5760
      TabIndex        =   31
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   30
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   29
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   28
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3840
      TabIndex        =   27
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      Height          =   495
      Index           =   4
      Left            =   8160
      TabIndex        =   26
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      Height          =   495
      Index           =   3
      Left            =   3360
      TabIndex        =   25
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   7680
      TabIndex        =   24
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   7200
      TabIndex        =   23
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   6720
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6240
      TabIndex        =   21
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5760
      TabIndex        =   20
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   19
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   18
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3840
      TabIndex        =   16
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   7680
      TabIndex        =   15
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   7200
      TabIndex        =   14
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   13
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      Height          =   495
      Index           =   22
      Left            =   8160
      TabIndex        =   12
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6480
      TabIndex        =   11
      Top             =   4320
      Width           =   75
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3840
      TabIndex        =   10
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      Height          =   495
      Index           =   2
      Left            =   3360
      TabIndex        =   9
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lABEL2 
      BackColor       =   &H00008000&
      Height          =   495
      Index           =   1
      Left            =   8640
      TabIndex        =   8
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   6720
      TabIndex        =   7
      Top             =   2640
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004000&
      BorderWidth     =   4
      X1              =   0
      X2              =   11880
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   6240
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   5760
      TabIndex        =   5
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   3
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4800
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   2640
      Top             =   1920
      Width           =   6615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004000&
      BackStyle       =   1  'Opaque
      Height          =   3975
      Left            =   2160
      Shape           =   2  'Oval
      Top             =   840
      Width           =   7815
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   4455
      Left            =   1920
      Top             =   720
      Width           =   8295
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu how 
         Caption         =   "How to..."
      End
      Begin VB.Menu about 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fad As String
Public m As String
Public dose As String
Public ans As String
Public f As Integer





Private Sub about_Click()
frmAbout.Visible = True
End Sub



Private Sub Command1_Click()
On Error Resume Next
off = RGB(205, 210, 200)
If Command3.Enabled = False Then
If Len(Text1.Text) = 1 Then
Call bono
Exit Sub
End If
End If
If Text1.Text = "" Then Exit Sub
If Len(Text1.Text) = 1 Then
If Command2.Enabled = True Then
Command1.Enabled = False
Else
Command2.Enabled = True
End If
End If
green = &H8000&
Command2.Enabled = True
Command3.Enabled = True
lblMsg1.Caption = "After guessing a letter, you must click Spin."
Text1.Text = UCase(Text1.Text)
Dim first As Integer, second As Integer, third As Integer, fourth As Integer
Dim wo, woo, wooo
If Len(Text1.Text) = 1 Then
lblLetters.Caption = lblLetters.Caption & Text1.Text & ", "
End If
Static f As Integer
Static m As String
Static fad As String
Static dose As String
Static A As Integer
f = f + 1
If f > 1 Then GoTo hell
If f < 2 Then
Randomize
A = Int((Val("26") * Rnd()) + 1)


Select Case A
Case 1:
Label3.Caption = "MOVIE"
fad = "THE"
m = "WIZARD"
dose = "OF OZ"
ans = "THE WIZARD OF OZ"
Case 2:
Label3.Caption = "ACTOR"
fad = "CLINT"
m = "EASTWOOD"
dose = ""
ans = "CLINT EASTWOOD"
Case 3:
Label3.Caption = "MOVIE"
fad = "THE"
m = "STAND"
dose = ""
ans = "THE STAND"
Case 4:
Label3.Caption = "BOOK"
fad = "THE CAT"
m = "IN THE"
dose = "  HAT"
ans = "THE CAT IN THE HAT"
Case 5:
Label3.Caption = "MOVIE"
fad = "HOME"
m = "ALONE"
dose = ""
ans = "HOME ALONE"
Case 6:
Label3.Caption = "SPACE"
fad = "THE"
m = "SOLAR"
dose = "SYSTEM"
ans = "THE SOLAR SYSTEM"
Case 7:
Label3.Caption = "PLANET"
m = "JUPITER"
dose = ""
ans = "JUPITER"
Case 8:
Label3.Caption = "MOVIE"
fad = "ACE"
m = "VENTURA"
dose = ""
ans = "ACE VENTURA"
Case 9:
Label3.Caption = "SCIENCE"
fad = "ALBERT"
m = "EINSTEIN"
dose = ""
ans = "ALBERT EINSTEIN"
Case 10:
Label3.Caption = "STATE"
fad = "NEW"
m = "YORK"
dose = ""
ans = "NEW YORK"
Case 11:
Label3.Caption = "PRESIDENT"
fad = "ABE"
m = "LINCOLN"
dose = ""
ans = "ABE LINCOLN"
Case 12:
Label3.Caption = "DRINK BRAND"
fad = "COCA COLA"
m = "CLASSIC"
dose = ""
ans = "COCA COLA CLASSIC"
Case 13:
Label3.Caption = "DRINK BRAND"
fad = "PEPSI"
m = "COLA"
dose = ""
ans = "PEPSI COLA"
Case 14:
Label3.Caption = "DRINK"
fad = "BOTTLED"
m = "WATER"
dose = ""
ans = "BOTTLED WATER"
Case 15:
Label3.Caption = "MUSIC TYPE"
fad = "PUNK"
m = "ROCK"
dose = ""
ans = "PUNK ROCK"
Case 16:
Label3.Caption = "STATE"
fad = "WEST"
m = "VIRGINIA"
dose = ""
ans = "WEST VIRGINIA"
Case 17:
Label3.Caption = "COUNTRY"
fad = "GREAT"
m = "BRITAIN"
dose = ""
ans = "GREAT BRITAIN"
Case 18:
Label3.Caption = "COUNTRY"
fad = "UNITED"
m = "STATES OF"
dose = "AMERICA"
ans = "UNITED STATES OF AMERICA"
Case 19:
Label3.Caption = "STATE"
fad = ""
m = "ILLINOIS"
dose = ""
ans = "ILLINOIS"
Case 20:
Label3.Caption = "STATE"
fad = "NEW"
m = "MEXICO"
dose = ""
ans = "NEW MEXICO"
Case 21:
Label3.Caption = "STATE"
fad = "SOUTH"
m = "DAKOTA"
dose = ""
ans = "SOUTH DAKOTA"
Case 22:
Label3.Caption = "STATE"
fad = "NORTH"
m = "CAROLINA"
dose = ""
ans = "NORTH CAROLINA"
Case 23:
Label3.Caption = "COUNRTY"
fad = "SOUTH"
m = "KOREA"
dose = ""
ans = "SOUTH KOREA"
Case 24:
Label3.Caption = "PRESIDENT"
fad = "BILL"
m = "CLINTON"
dose = ""
ans = "BILL CLINTON"
Case 25:
Label3.Caption = "CITY"
fad = "LOS"
m = "ANGELOS"
dose = ""
ans = "LOS ANGELOS"

Case 26:
Label3.Caption = "CONTINENT"
fad = "NORTH"
m = "AMERICA"
dose = ""
ans = "NORTH AMERICA"
Case 27:
Label3.Caption = "STATE"
fad = "SOUTH"
m = "CAROLINA"
ans = ""

End Select

End If

hell:
If Len(m) = 3 Then
l = Left(m, 1)
s = Mid(m, 2, 1)
ff = Right(m, 1)


For TOL = 0 To 2
If Label1(TOL).BackColor = vbWhite Then
End If
If Label1(TOL).BackColor <> vbWhite Then
Label1(TOL).BackColor = off
End If
Next
End If

If Len(m) = 4 Then
l = Left(m, 1)
s = Mid(m, 2, 1)
ff = Mid(m, 3, 1)
r = Right(m, 1)



For TOL = 0 To 3
If Label1(TOL).BackColor = vbWhite Then
End If
If Label1(TOL).BackColor <> vbWhite Then
Label1(TOL).BackColor = off
End If
Next
End If



If Len(m) = 5 Then
l = Left(m, 1)
s = Mid(m, 2, 1)
ff = Mid(m, 3, 1)
r = Mid(m, 4, 1)
i = Right(m, 1)
For TOL = 0 To 4
If Label1(TOL).BackColor = vbWhite Then
End If
If Label1(TOL).BackColor <> vbWhite Then
Label1(TOL).BackColor = off
End If
Next
End If

If Len(m) = 6 Then
l = Left(m, 1)
s = Mid(m, 2, 1)
ff = Mid(m, 3, 1)
r = Mid(m, 4, 1)
i = Mid(m, 5, 1)
b = Right(m, 1)
For TOL = 0 To 5
If Label1(TOL).BackColor = vbWhite Then
End If
If Label1(TOL).BackColor <> vbWhite Then
Label1(TOL).BackColor = off
End If
Next
End If


If Len(m) = 7 Then
l = Left(m, 1)
s = Mid(m, 2, 1)
ff = Mid(m, 3, 1)
r = Mid(m, 4, 1)
i = Mid(m, 5, 1)
b = Mid(m, 6, 1)
h = Right(m, 1)
For TOL = 0 To 6
If Label1(TOL).BackColor = vbWhite Then
End If
If Label1(TOL).BackColor <> vbWhite Then
Label1(TOL).BackColor = off
End If
Next
End If


If Len(m) = 8 Then
l = Left(m, 1)
s = Mid(m, 2, 1)
ff = Mid(m, 3, 1)
r = Mid(m, 4, 1)
i = Mid(m, 5, 1)
b = Mid(m, 6, 1)
h = Mid(m, 7, 1)
bo = Right(m, 1)
For TOL = 0 To 7
If Label1(TOL).BackColor = vbWhite Then
End If
If Label1(TOL).BackColor <> vbWhite Then
Label1(TOL).BackColor = off
End If
Next
End If

If Len(m) = 9 Then
l = Left(m, 1)
s = Mid(m, 2, 1)
ff = Mid(m, 3, 1)
r = Mid(m, 4, 1)
i = Mid(m, 5, 1)
b = Mid(m, 6, 1)
h = Mid(m, 7, 1)
bo = Mid(m, 8, 1)
ca = Right(m, 1)
For TOL = 0 To 8
If Label1(TOL).BackColor = vbWhite Then
End If
If Label1(TOL).BackColor <> vbWhite Then
Label1(TOL).BackColor = off
End If
Next
ass.Caption = ass.Caption + Label7.Caption
End If

'this for top line
'PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP
'###############################################################
'#@#$#########################################################

If Len(fad) = 3 Then
one = Left(fad, 1)
two = Mid(fad, 2, 1)
three = Right(fad, 1)
For TOL = 0 To 2
If Label4(TOL).BackColor = vbWhite Then
End If
If Label4(TOL).BackColor <> vbWhite Then
Label4(TOL).BackColor = off
End If
Next
End If

If Len(fad) = 4 Then
one = Left(fad, 1)
two = Mid(fad, 2, 1)
three = Mid(fad, 3, 1)
four = Right(fad, 1)
For TOL = 0 To 3
If Label4(TOL).BackColor = vbWhite Then
End If
If Label4(TOL).BackColor <> vbWhite Then
Label4(TOL).BackColor = off
End If
Next
End If

If Len(fad) = 5 Then
one = Left(fad, 1)
two = Mid(fad, 2, 1)
three = Mid(fad, 3, 1)
four = Mid(fad, 4, 1)
five = Right(fad, 1)
For TOL = 0 To 4
If Label4(TOL).BackColor = vbWhite Then
End If
If Label4(TOL).BackColor <> vbWhite Then
Label4(TOL).BackColor = off
End If

Next
End If


If Len(fad) = 6 Then
one = Left(fad, 1)
two = Mid(fad, 2, 1)
three = Mid(fad, 3, 1)
four = Mid(fad, 4, 1)
five = Mid(fad, 5, 1)
six = Right(fad, 1)
For TOL = 0 To 5
If Label4(TOL).BackColor = vbWhite Then
End If
If Label4(TOL).BackColor <> vbWhite Then
Label4(TOL).BackColor = off
End If
Next
End If

If Len(fad) = 7 Then
one = Left(fad, 1)
two = Mid(fad, 2, 1)
three = Mid(fad, 3, 1)
four = Mid(fad, 4, 1)
five = Mid(fad, 5, 1)
six = Mid(fad, 6, 1)
seven = Right(fad, 1)
For TOL = 0 To 6
If Label4(TOL).BackColor = vbWhite Then
End If
If Label4(TOL).BackColor <> vbWhite Then
Label4(TOL).BackColor = off
End If
Next
End If

If Len(fad) = 8 Then
one = Left(fad, 1)
two = Mid(fad, 2, 1)
three = Mid(fad, 3, 1)
four = Mid(fad, 4, 1)
five = Mid(fad, 5, 1)
six = Mid(fad, 6, 1)
seven = Mid(fad, 7, 1)
eight = Right(fad, 1)
For TOL = 0 To 7
If Label4(TOL).BackColor = vbWhite Then
End If
If Label4(TOL).BackColor <> vbWhite Then
Label4(TOL).BackColor = off
End If
Next
End If

If Len(fad) = 9 Then
one = Left(fad, 1)
two = Mid(fad, 2, 1)
three = Mid(fad, 3, 1)
four = Mid(fad, 4, 1)
five = Mid(fad, 5, 1)
six = Mid(fad, 6, 1)
seven = Mid(fad, 7, 1)
eight = Mid(fad, 8, 1)
nine = Right(fad, 1)
For TOL = 0 To 8
If Label4(TOL).BackColor = vbWhite Then
End If
If Label4(TOL).BackColor <> vbWhite Then
Label4(TOL).BackColor = off
End If
Next
End If

'THIS IS THE BOTTOM LINE
'LAST ROW ON BOTTOM#######################################################
'*********************************************************************


If Len(dose) = 3 Then
os1 = Left(dose, 1)
os2 = Mid(dose, 2, 1)
os3 = Right(dose, 1)
For TOL = 0 To 2
If Label5(TOL).BackColor = vbWhite Then
End If
If Label5(TOL).BackColor <> vbWhite Then
Label5(TOL).BackColor = off
End If
Next
End If

If Len(dose) = 4 Then
os1 = Left(dose, 1)
os2 = Mid(dose, 2, 1)
os3 = Mid(dose, 3, 1)
os4 = Right(dose, 1)
For TOL = 0 To 3
If Label5(TOL).BackColor = vbWhite Then
End If
If Label5(TOL).BackColor <> vbWhite Then
Label5(TOL).BackColor = off
End If
Next
End If

If Len(dose) = 5 Then
os1 = Left(dose, 1)
os2 = Mid(dose, 2, 1)
os3 = Mid(dose, 3, 1)
os4 = Mid(dose, 4, 1)
os5 = Right(dose, 1)
For TOL = 0 To 4
If Label5(TOL).BackColor = vbWhite Then
End If
If Label5(TOL).BackColor <> vbWhite Then
Label5(TOL).BackColor = off
End If
Next
End If

If Len(dose) = 6 Then
os1 = Left(dose, 1)
os2 = Mid(dose, 2, 1)
os3 = Mid(dose, 3, 1)
os4 = Mid(dose, 4, 1)
os5 = Mid(dose, 5, 1)
os6 = Right(dose, 1)
For TOL = 0 To 5
If Label5(TOL).BackColor = vbWhite Then
End If
If Label5(TOL).BackColor <> vbWhite Then
Label5(TOL).BackColor = off
End If
Next
End If

If Len(dose) = 7 Then
os1 = Left(dose, 1)
os2 = Mid(dose, 2, 1)
os3 = Mid(dose, 3, 1)
os4 = Mid(dose, 4, 1)
os5 = Mid(dose, 5, 1)
os6 = Mid(dose, 6, 1)
os7 = Right(dose, 1)
For TOL = 0 To 6
If Label5(TOL).BackColor = vbWhite Then
End If
If Label5(TOL).BackColor <> vbWhite Then
Label5(TOL).BackColor = off
End If
Next
End If

If Len(dose) = 8 Then
os1 = Left(dose, 1)
os2 = Mid(dose, 2, 1)
dos3 = Mid(dose, 3, 1)
os4 = Mid(dose, 4, 1)
os5 = Mid(dose, 5, 1)
os6 = Mid(dose, 6, 1)
os7 = Mid(dose, 7, 1)
os8 = Right(dose, 1)
For TOL = 0 To 7
If Label5(TOL).BackColor = vbWhite Then
End If
If Label5(TOL).BackColor <> vbWhite Then
Label5(TOL).BackColor = off
End If
Next
End If

If Len(dose) = 9 Then
os1 = Left(dose, 1)
os2 = Mid(dose, 2, 1)
os3 = Mid(dose, 3, 1)
os4 = Mid(dose, 4, 1)
os5 = Mid(dose, 5, 1)
os6 = Mid(dose, 6, 1)
os7 = Mid(dose, 7, 1)
os8 = Mid(dose, 8, 1)
os9 = Right(dose, 1)
For TOL = 0 To 8
If Label5(TOL).BackColor = vbWhite Then
End If
If Label5(TOL).BackColor <> vbWhite Then
Label5(TOL).BackColor = off
End If
Next
End If


'#########################################################################
'##################################################################
'AFTER YOU GET SOUND WORKING TAKE OUT MSGBOX "BEEP" AND PUT BEEP
'###################################################################
'#############################################################################
If Text1.Text = "POINTS" Then
ass.Caption = Int(99999)
End If
If ass.Caption > Int(10000000) Then
End
End If
If Text1.Text = "RASBO" Then
MsgBox "RASBO,FUNKY BUTT CRUNCH NATO ACTION BONER PIE SANCHO STYLE, WITH A CAMO RAMBO QWEEEFER IN THE TRENCHES. NATO CHEMO WARFARE WITH THE TRENCHCOAT TWINS AND A BUNCH OF FUNKY BOO BOO RAUNHY CAMEL BLISTERS ON YOUR I CANT BELIEVE ITS NOT BUTTERNUT CRUNCHER"
End If
If Text1.Text = l Then

Label1(0).Caption = l
Label1(0).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
 Beep
End If
End If
If Text1.Text = s Then
Label1(1).Caption = s
Label1(1).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = ff Then
Label1(2).Caption = ff
Label1(2).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
 Beep
End If
End If
If Text1.Text = r Then
Label1(3).Caption = r
Label1(3).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = t Then
Label1(4).Caption = t
Label1(4).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = i Then
Label1(4).Caption = i
Label1(4).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = b Then
Label1(5).Caption = b
Label1(5).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = h Then
Label1(6).Caption = h
Label1(6).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = bo Then
Label1(7).Caption = bo
Label1(7).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = ca Then
Label1(8).Caption = ca
Label1(8).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
'fad is here
'aadfafadfdafafadfadfafad'adgfad'
If Text1.Text = one Then
Label4(0).Caption = one
Label4(0).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = two Then
Label4(1).Caption = two
Label4(1).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = three Then
Label4(2).Caption = three
Label4(2).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = four Then
Label4(3).Caption = four
Label4(3).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = five Then
Label4(4).Caption = five
Label4(4).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = six Then
Label4(5).Caption = six
Label4(5).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = seven Then
Label4(6).Caption = seven
Label4(6).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = eight Then
Label4(7).Caption = eight
Label4(7).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = nine Then
Label4(8).Caption = nine
Label4(8).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os1 Then
Label5(0).Caption = os1
Label5(0).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os2 Then
Label5(1).Caption = os2
Label5(1).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os3 Then
Label5(2).Caption = os3
Label5(2).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os4 Then
Label5(3).Caption = os4
Label5(3).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os5 Then
Label5(4).Caption = os5
Label5(4).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os6 Then
Label5(5).Caption = os6
Label5(5).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os7 Then
Label5(6).Caption = os7
Label5(6).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os8 Then
Label5(7).Caption = os8
Label5(7).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If
If Text1.Text = os9 Then
Label5(8).Caption = os9
Label5(8).BackColor = vbWhite
ass.Caption = Int(ass.Caption) + Int(Label7.Caption)
If frmStart1.Label4.Caption = "ON" Then
Beep
End If
End If



If Len(Text1.Text) > 1 Then
    If Timer1.Enabled = True Then
    If Text1.Text = ans Then
    Text1.Text = ""
    MsgBox "You have correctly guessed the final puzzle!"
    MsgBox "Please play again soon!", vbOKOnly, "Wheel-Of-Fortune"
    End
    Exit Sub
    End If
    End If
    If Text1.Text = ans Then
    MsgBox " " & ans & " is the correct answer, " & frmStart1.Text1.Text & ". You recieved a $1000 bonus!"
    Text1.Text = ""
    ass.Caption = Int(ass.Caption) + Int(1000)
    If Int(ass.Caption) > Int((frmStart1.level1.Caption) * 4000) Then
    Text1.Enabled = True
    For dad = 0 To 8
    If Label4(dad).BackColor = vbWhite Or Label4(dad).BackColor = off Then
    Label4(dad).BackColor = green
    Label4(dad).Caption = ""
    End If
    Next
    For dadd = 0 To 8
    If Label1(dadd).BackColor = vbWhite Or Label1(dadd).BackColor = off Then
    Label1(dadd).BackColor = green
    Label1(dadd).Caption = ""
    Else
    Label4(dadd).BackColor = green
    Label4(dadd).Caption = ""
    End If
    Next
    For dada = 0 To 8
    If Label5(dada).BackColor = vbWhite Or Label5(dada).BackColor = off Then
    Label5(dada).BackColor = green
    Label5(dada).Caption = ""
    End If
    Next
    m = ""
    fad = ""
    dose = ""
    Randomize
    A = Int((Val("26") * Rnd) + 1)
    Label3.Caption = ""
    lblLetters.Caption = ""
    Label7.Caption = Int(0)
    f = 0
    Call final
    Else
    Text1.Text = ""
    ' exit sub, end if
    lblMsg1.Caption = "To start next puzzle,please spin."
    Command3.Enabled = True
    For dad = 0 To 8
    If Label4(dad).BackColor = vbWhite Or Label4(dad).BackColor = off Then
    Label4(dad).BackColor = green
    Label4(dad).Caption = ""
    End If
    Next
    For dadd = 0 To 8
    If Label1(dadd).BackColor = vbWhite Or Label1(dadd).BackColor = off Then
    Label1(dadd).BackColor = green
    Label1(dadd).Caption = ""
    Else
    Label4(dadd).BackColor = green
    Label4(dadd).Caption = ""
    End If
    Next
    For dada = 0 To 8
    If Label5(dada).BackColor = vbWhite Or Label5(dada).BackColor = off Then
    Label5(dada).BackColor = green
    Label5(dada).Caption = ""
    End If
    Next
    m = ""
    fad = ""
    dose = ""
    Randomize
    A = Int((Val("26") * Rnd) + 1)
    Label3.Caption = ""
    lblLetters.Caption = ""
    Label7.Caption = Int(0)
    f = 0
    Exit Sub
    End If
Else
MsgBox "That was an incorrect answer, " & frmStart1.Text1.Text & ". Please choose another letter."
Command2.Enabled = False
Command3.Enabled = True
Text1.Enabled = True
lblMsg1.Caption = "Please select another letter to guess."
Text1.SetFocus
End If
End If

If four = " " Then MsgBox "The cat in the hat"




Text1.Text = ""


End Sub



Private Sub Command2_Click()
Text1.Enabled = True
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Dim gate As Integer
Dim bill As Integer
gate = Int(Val("10") * Rnd)
bill = Int(gate * 100)
Label7.Caption = bill
lblMsg1.Caption = "After spinning, you must guess another letter or solve the puzzle."
Text1.SetFocus
End Sub

Private Sub Command3_Click()
If Command2.Enabled = True Then
Command2.Enabled = False
Command3.Enabled = False
Text1.Enabled = True
End If
Form11.lblMsg1.Caption = "Enter the answer then hit ok."
End Sub




Private Sub exit_Click()
MsgBox "Thanks for playing Wheel-Of-Fortune, " & frmStart1.Text1.Text & ". Hope you will play again."
Form11.Visible = False
Sosa.Visible = True
End Sub

Private Sub Form_Load()
If frmStart1.Text2.Visible = True Then
Label10.Visible = True
Label10.Caption = frmStart1.Text2.Text
Label11.Visible = True
Else
Label10.Visible = False
Label11.Visible = False
End If
Label8.Caption = frmStart1.Text1.Text
Label7.Caption = Val("100")
Command1.Enabled = False
Text1.ToolTipText = "Enter one letter to guess the puzzle. Enter the answer to the puzzle if you know it."
Command1.ToolTipText = "Click this to submit a guess at the puzzle or after you entered a possible answer to the puzzle."
Command2.ToolTipText = "Click this after you attempted a guess to spin for a money value"
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub how_Click()
frmHelp.Visible = True
End Sub

Private Sub Text1_Change()
If Command2.Enabled = True Then
Text1.Enabled = False
Else
Text1.Enabled = True
End If
If Len(Text1.Text) > 0 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub

Sub bono()
MsgBox "One letter answers are invalid when trying to solve the puzzle. Please select again.", vbCritical + vbOKOnly, "Hold on...."
End Sub

Sub final()
green = &H8000&
lblMsg1.Caption = "You have made it to the final round. Click Ok to begin"
For dad = 0 To 8
    If Label4(dad).BackColor = vbWhite Then
    Label4(dad).BackColor = green
    Label4(dad).Caption = ""
    End If
    Next
    For dadd = 0 To 8
    If Label1(dadd).BackColor = vbWhite Then
    Label1(dadd).BackColor = green
    Label1(dadd).Caption = ""
    Else
    Label4(dadd).BackColor = green
    Label4(dadd).Caption = ""
    End If
    Next
    For dada = 0 To 8
    If Label5(dada).BackColor = vbWhite Then
    Label5(dada).BackColor = green
    Label5(dada).Caption = ""
    End If
    Next
     m = ""
    fad = ""
    dose = ""
    Randomize
    A = Int((Val("26") * Rnd) + 1)
    Label3.Caption = ""
    lblLetters.Caption = ""
    Label7.Caption = Int(0)
    f = 0
Label9.Visible = True
Text1.Text = ""
Text1.Enabled = True
Text1.SetFocus
MsgBox "Click ok to begin. You will only have thirty(30) seconds to solve the puzzle. Good luck.", vbOKOnly, "Final puzzle..."
lblMsg1.Caption = "You have thirty seconds to solve this puzzle."
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If Label9.Caption < 2 Then Call endit: Exit Sub
Label9.Caption = Int(Label9.Caption) - Int(1)
End Sub


Sub endit()
MsgBox "You have finished the game with a score of " & ass.Caption
MsgBox "Please play Wheel-Of-Fortune again,and soon!", App.Title
End
End Sub
