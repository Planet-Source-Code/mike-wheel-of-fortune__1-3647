VERSION 5.00
Begin VB.Form frmSpin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   3720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spin"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   3720
      Width           =   615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000C000&
      BorderWidth     =   12
      X1              =   3360
      X2              =   3360
      Y1              =   240
      Y2              =   2520
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF80FF&
      BorderWidth     =   12
      X1              =   2160
      X2              =   4440
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   12
      X1              =   2400
      X2              =   4200
      Y1              =   2040
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000040C0&
      BorderWidth     =   12
      X1              =   2520
      X2              =   4200
      Y1              =   600
      Y2              =   2040
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   1200
      Shape           =   3  'Circle
      Top             =   240
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C000&
      BorderWidth     =   3
      Height          =   1455
      Left            =   720
      Top             =   2760
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Type a spin speed between 1 and 10."
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   4095
   End
End
Attribute VB_Name = "frmSpin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Do
If Line1.BorderColor = &H40C0& Then
Line1.BorderColor = vbBlue
Else
Line1.BorderColor = &H40C0&
End If
Loop


End Sub
