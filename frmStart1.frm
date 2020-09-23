VERSION 5.00
Begin VB.Form frmStart1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wheel-Of-Fortune"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Skill level"
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   2400
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
      Begin VB.HScrollBar HScroll3 
         Height          =   255
         Left            =   480
         Max             =   10
         Min             =   1
         TabIndex        =   17
         Top             =   840
         Value           =   1
         Width           =   375
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "(Easy)"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.Label level1 
         BackColor       =   &H00000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS LineDraw"
            Size            =   14.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Players"
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "NewBskvll BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   4200
      Width           =   3135
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Sound"
      ForeColor       =   &H0000FF00&
      Height          =   1215
      Left            =   3960
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "OFF"
         BeginProperty Font 
            Name            =   "NewBskvll BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00008000&
      Height          =   1695
      Left            =   480
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   5
      Left            =   4680
      Top             =   600
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   4
      Left            =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   3
      Left            =   4680
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   2
      Left            =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   4680
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   0
      Left            =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   7
      X1              =   1440
      X2              =   4680
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   6
      X1              =   1560
      X2              =   4560
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   5
      X1              =   4560
      X2              =   4560
      Y1              =   360
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   4
      X1              =   1440
      X2              =   1440
      Y1              =   240
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   3
      X1              =   4680
      X2              =   4680
      Y1              =   240
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   2
      X1              =   1440
      X2              =   4680
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   1
      X1              =   1560
      X2              =   4560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Index           =   0
      X1              =   1560
      X2              =   1560
      Y1              =   360
      Y2              =   2160
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "FORTUNE"
      BeginProperty Font 
         Name            =   "AdLib BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   2040
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "of"
      BeginProperty Font 
         Name            =   "AdLib BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2640
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Player 1:"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lblWheel 
      BackColor       =   &H00000000&
      Caption         =   "WHEEL"
      BeginProperty Font 
         Name            =   "AdLib BT"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "frmStart1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlay_Click()
If Text2.Visible = False Then
If Text1.Text = "" Then
MsgBox "enter"
End If
End If
If Text2.Visible = False Then
If Text1.Text <> "" Then
Form11.Visible = True
frmStart1.Visible = False
End If
End If
If Text2.Visible = True Then
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "nope"
Else
Form11.Visible = True
frmStart1.Visible = False
End If
End If
End Sub


Private Sub Command1_Click()
Sosa.Visible = True
frmStart1.Visible = False
End Sub

Private Sub Form_Load()
cmdPlay.Enabled = False
End Sub

Private Sub HScroll1_Change()
If Label4.Caption = "OFF" Then
Label4.Caption = "ON"
Else
Label4.Caption = "OFF"
End If
End Sub

Private Sub HScroll2_Change()
If Label5.Caption = "1" Then
Label5.Caption = "2"
Text2.Visible = True
cmdPlay.Enabled = False
Label6.Visible = True
Else
Label5.Caption = "1"
Text2.Visible = False
Text2.Text = ""
If Text1.Text <> "" Then
cmdPlay.Enabled = True
End If
Text1.SetFocus
Label6.Visible = False
End If
End Sub

Private Sub HScroll3_Change()
level1.Caption = HScroll3.Value
If level1.Caption = 1 Then
Label7.Caption = "(Easy)"
End If
If level1.Caption = 2 Then
Label7.Caption = "(Easy)"
End If
If level1.Caption = 3 Then
Label7.Caption = "(Easy)"
End If
If level1.Caption = 4 Then
Label7.Caption = "(Med.)"
End If
If level1.Caption = 5 Then
Label7.Caption = "(Med.)"
End If
If level1.Caption = 6 Then
Label7.Caption = "(Med.)"
End If
If level1.Caption = 7 Then
Label7.Caption = "(Med.)"
End If
If level1.Caption = 8 Then
Label7.Caption = "(Hard)"
End If
If level1.Caption = 9 Then
Label7.Caption = "(Hard)"
End If
If level1.Caption = 10 Then
Label7.Caption = "(Hard)"
End If





End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
cmdPlay.Enabled = False
End If
If Text2.Visible = False Then
If Text1.Text = "" Then
cmdPlay.Enabled = False
Else
cmdPlay.Enabled = True
End If
End If
If Text2.Visible = True Then
If Text1.Text = "" Then
cmdPlay.Enabled = False
End If
If Text2.Text <> "" Then
If Text1.Text <> "" Then
cmdPlay.Enabled = True
End If
Else
cmdPlay.Enabled = False
End If
End If
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) = 0 Then
cmdPlay.Enabled = False
End If
If Text1.Text <> "" Then
If Len(Text2.Text) >= 1 Then
cmdPlay.Enabled = True
Else
cmdPlay.Enabled = False
End If
End If
End Sub
