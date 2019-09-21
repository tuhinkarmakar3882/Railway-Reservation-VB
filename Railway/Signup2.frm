VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Signup2 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form3"
   ClientHeight    =   10215
   ClientLeft      =   2235
   ClientTop       =   465
   ClientWidth     =   15165
   LinkTopic       =   "Form3"
   ScaleHeight     =   10215
   ScaleWidth      =   15165
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Signup2.frx":0000
      Left            =   3000
      List            =   "Signup2.frx":000A
      TabIndex        =   23
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Signup2.frx":0022
      Left            =   3000
      List            =   "Signup2.frx":002F
      TabIndex        =   21
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   20
      Top             =   9480
      Width           =   1095
   End
   Begin VB.CommandButton next2 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   19
      Top             =   9480
      Width           =   1095
   End
   Begin VB.ComboBox work 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Signup2.frx":0048
      Left            =   3000
      List            =   "Signup2.frx":0058
      TabIndex        =   18
      Text            =   "Select your Occupation"
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   17
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   16
      Top             =   7560
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      MaxLength       =   16
      TabIndex        =   15
      Top             =   6240
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      MaxLength       =   12
      TabIndex        =   14
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   13
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   12
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   11
      Top             =   960
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   22
      Top             =   4320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   146931713
      CurrentDate     =   40151
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Details :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label mobile 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile *"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   10
      Top             =   8520
      Width           =   1095
   End
   Begin VB.Label email 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail *"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label pan 
      BackStyle       =   0  'Transparent
      Caption         =   "Pan Card No"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   8
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label adhar 
      BackStyle       =   0  'Transparent
      Caption         =   "Aadhard Card No*"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   7
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label occupation 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation *"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   6
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label dob 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth *"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label marry 
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label gender 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender *"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label name_last 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name*"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label name_mid 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label name_first 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name *"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   500
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Signup2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Signup1.Show

End Sub

Private Sub nation_Change()

End Sub

Private Sub next2_Click()
If Text1.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or DTPicker1.Value = "" Or work = "" Or Text6.Text = "" Or Text5.Text = "" Or Text8.Text = "" Then
MsgBox ("Please Enter the star(*)  Marked Fields.")
Else
Signup3.Show
Me.Hide
End If

End Sub
