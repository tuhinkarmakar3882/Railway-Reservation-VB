VERSION 5.00
Begin VB.Form Signup3 
   BackColor       =   &H00C0FFC0&
   ClientHeight    =   10215
   ClientLeft      =   2445
   ClientTop       =   675
   ClientWidth     =   14595
   LinkTopic       =   "Form4"
   ScaleHeight     =   10215
   ScaleWidth      =   14595
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Signup3.frx":0000
      Left            =   6000
      List            =   "Signup3.frx":000A
      TabIndex        =   26
      Top             =   8160
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Signup3.frx":0017
      Left            =   6000
      List            =   "Signup3.frx":0021
      TabIndex        =   25
      Top             =   7560
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Signup3.frx":002E
      Left            =   6000
      List            =   "Signup3.frx":0038
      TabIndex        =   24
      Top             =   6960
      Width           =   1335
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
      Height          =   375
      Left            =   6240
      TabIndex        =   23
      Top             =   9480
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6240
      TabIndex        =   22
      Top             =   5160
      Width           =   735
   End
   Begin VB.OptionButton yes 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4920
      TabIndex        =   21
      Top             =   5160
      Width           =   735
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
      Left            =   4800
      TabIndex        =   20
      Top             =   4440
      Width           =   2175
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
      Left            =   4800
      TabIndex        =   19
      Top             =   3840
      Width           =   2175
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
      Left            =   4800
      TabIndex        =   18
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox Text4 
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
      Left            =   4800
      TabIndex        =   17
      Top             =   2640
      Width           =   2175
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
      Left            =   4800
      TabIndex        =   16
      Top             =   2040
      Width           =   2175
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
      Left            =   4800
      TabIndex        =   15
      Top             =   1440
      Width           =   2175
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
      Left            =   4800
      TabIndex        =   14
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton submit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subscription to IRCTC Newsletters/Promotions/IRCTC SBI Card through email/SMS :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   6480
      Width           =   6975
   End
   Begin VB.Label inform 
      BackStyle       =   0  'Transparent
      Caption         =   "Please inform me about IRCTC,SBI card through phone *"
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
      Left            =   480
      TabIndex        =   11
      Top             =   8280
      Width           =   4695
   End
   Begin VB.Label promo 
      BackStyle       =   0  'Transparent
      Caption         =   "Commercial Promotions"
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
      Left            =   480
      TabIndex        =   10
      Top             =   7680
      Width           =   2055
   End
   Begin VB.Label newsletters 
      BackStyle       =   0  'Transparent
      Caption         =   "IRCTC Newsletters"
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
      Left            =   480
      TabIndex        =   9
      Top             =   7080
      Width           =   1815
   End
   Begin VB.Label copy 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy residence to office address"
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
      Left            =   480
      TabIndex        =   8
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label no 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Office *"
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
      Left            =   480
      TabIndex        =   7
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label city 
      BackStyle       =   0  'Transparent
      Caption         =   "City/Town *"
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
      Left            =   480
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label state 
      BackStyle       =   0  'Transparent
      Caption         =   "State *"
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
      Left            =   480
      TabIndex        =   5
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label pin 
      BackStyle       =   0  'Transparent
      Caption         =   "Pincode *"
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
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label area 
      BackStyle       =   0  'Transparent
      Caption         =   "Area/Locality"
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
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Lane 
      BackStyle       =   0  'Transparent
      Caption         =   "Street/Lane"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label block 
      BackStyle       =   0  'Transparent
      Caption         =   "Flat/Door/Block no *"
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
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label addr 
      BackStyle       =   0  'Transparent
      Caption         =   "Resedential Address :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Signup3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Signup2.Show

End Sub

Private Sub submit_Click()
If Text1.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Combo3.Text = "" Then
MsgBox ("Please Enter the star(*)  Marked Fields.")
Else
MsgBox ("Congrats! You've Successfully Created Your Account!!")
Me.Hide
Form14.Show
End If
End Sub

