VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000014&
   Caption         =   "How-To Play "
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8205
   LinkTopic       =   "Form2"
   ScaleHeight     =   6015
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   855
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmHelp.frx":0000
      Top             =   3840
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return to Wheel-Of-Fortune"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   6
      Text            =   "Click the Spin buton to see how much money you'll get ."
      Top             =   3120
      Width           =   5055
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Spin"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Text            =   "If you know the answer to the puzzle,click Solve, and then type it out."
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Ok/Enter Letter"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "Enter one letter at a time if you want to keep guessing at the puzzle."
      Top             =   1440
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Ok/Enter Letter"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "ENTER "" POINTS"" AS YOUR ANSWER FOR MAXIMUM BONUS"
      Height          =   975
      Left            =   9000
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   120
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   120
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   7440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Wheel Of Fortune Help"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   6375
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmHelp.Visible = False

End Sub
