VERSION 5.00
Begin VB.Form Register2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register"
   ClientHeight    =   4230
   ClientLeft      =   3015
   ClientTop       =   2355
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6000
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      Top             =   450
      Width           =   4575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   2685
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Serial:"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Orgarnization:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   160
      Width           =   855
   End
End
Attribute VB_Name = "Register2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Clipboard.SetText (Text3.Text)
Unload Me
Register.Show
End Sub

Private Sub Command2_Click()
Unload Me
Register.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
cnt = Len(Text1.Text)
abc = KeyAscii
If abc <> 8 Then
If Len(Text3.Text) > 1330 Then
MsgBox "Overflow"
Text1.Text = ""
Text3.Text = ""
End If
Text3.SelText = (CInt(Sqr(abc * 20 - 40 + 160) * cnt))
Else
Text1.Text = ""
Text3.Text = ""
End If
End Sub
