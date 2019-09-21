VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   LinkTopic       =   "Form14"
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   1000
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ABOUT-US"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Width           =   1000
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "SIGN UP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to 2-R-S-T Railway Services"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   4320
      TabIndex        =   0
      Top             =   720
      Width           =   3270
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
frmSplash.Show

End Sub

Private Sub Command2_Click()
Me.Hide
Signup1.Show
End Sub

Private Sub Command3_Click()
Me.Hide
frmAbout.Show

End Sub

Private Sub Command4_Click()
Form15.Show
Me.Hide
End Sub

