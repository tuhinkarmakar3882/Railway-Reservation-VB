VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   4395
   ClientLeft      =   6870
   ClientTop       =   2970
   ClientWidth     =   4800
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   1000
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   1000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   1000
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   1000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "newfrm.frx":0000
      Left            =   2160
      List            =   "newfrm.frx":000A
      TabIndex        =   3
      Top             =   2200
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Food From.."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"newfrm.frx":0023
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tummy on Rails"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   4815
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox ("Please Choose A Distribution.")
Else
MsgBox ("Your order from " + Combo1.Text + " has been received. The food will be served at Mughal Sarai Station")
Domdom.Show
Me.Hide
End If
End Sub


Private Sub Command2_Click()
MsgBox ("Thank You. Do Visit Us Again.")
Unload Me
End Sub

