VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11475
   FillColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   Picture         =   "form pay.frx":0000
   ScaleHeight     =   7005
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   20
      Top             =   5520
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   8280
      MaxLength       =   4
      TabIndex        =   15
      Top             =   2000
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   300
      Left            =   5040
      MaxLength       =   4
      TabIndex        =   14
      Top             =   2000
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   13
      Top             =   2000
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   7200
      MaxLength       =   4
      TabIndex        =   12
      Top             =   2000
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "form pay.frx":13756
      Left            =   6240
      List            =   "form pay.frx":1377E
      TabIndex        =   11
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "PAY NOW"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      MaskColor       =   &H0000C0C0&
      TabIndex        =   10
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8520
      MaxLength       =   4
      TabIndex        =   7
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5040
      MaxLength       =   3
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2900
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER THE TEXT BELOW:"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   19
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Image Cap1 
      Height          =   795
      Left            =   5040
      Picture         =   "form pay.frx":137BF
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   1995
   End
   Begin VB.Image Cap3 
      Height          =   795
      Left            =   5040
      Picture         =   "form pay.frx":194F4
      Stretch         =   -1  'True
      Top             =   4500
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image Cap4 
      Height          =   795
      Left            =   5040
      Picture         =   "form pay.frx":2B98E
      Stretch         =   -1  'True
      Top             =   4500
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image Cap2 
      Height          =   795
      Left            =   5040
      Picture         =   "form pay.frx":31082
      Stretch         =   -1  'True
      Top             =   4500
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image Image3 
      Height          =   450
      Left            =   3735
      Picture         =   "form pay.frx":38595
      Stretch         =   -1  'True
      Top             =   2000
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   450
      Left            =   3735
      Picture         =   "form pay.frx":3A83F
      Stretch         =   -1  'True
      Top             =   2000
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   3735
      Picture         =   "form pay.frx":3B35B
      Stretch         =   -1  'True
      Top             =   2000
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8040
      TabIndex        =   18
      Top             =   2000
      Width           =   135
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   17
      Top             =   2000
      Width           =   135
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      TabIndex        =   16
      Top             =   2000
      Width           =   135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MONTH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER THE EXPIRY DATE"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   1200
      TabIndex        =   4
      Top             =   3600
      Width           =   2000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER CVV NUMBER"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2900
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "CARD HOLDER'S NAME"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CARD NUMBER"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2000
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT GATEWAY"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()

End Sub



Private Sub Command1_Click()
If Text2.Text = "" Or Len(Text6.Text) <> 4 Or Len(Text4.Text) <> 4 Or Len(Text1.Text) <> 4 Or Len(Text7.Text) <> 4 Or Len(Text3.Text) <> 3 Or Val(Text5.Text) < 2019 Or Combo1.Text = "" Then
MsgBox ("Please Enter Valid Details! Make Sure You Have Entered All the Fields.")
Else
MsgBox ("Payment Successful")
Form19.WindowsMediaPlayer1.Enabled = True
Form16.tst.Recordset.AddNew
Form16.Text1.Text = MDIForm1.Text1.Text

Me.Hide
Form19.Show
End If
End Sub

Private Sub Form_Load()

Image1.Visible = False
Image2.Visible = False
Image3.Visible = False

End Sub

Private Sub Text2_Lostfocus()
If (Len(Text2.Text) <> 0) Then
Text6.SetFocus
End If
End Sub





Private Sub Text7_Change()


If Len(Text7.Text) = 4 Then
    Text3.SetFocus
End If

End Sub


Private Sub Text1_Change()


If Len(Text1.Text) = 4 Then
    Text7.SetFocus
End If

End Sub

Private Sub Text4_Change()
If Len(Text4.Text) = 4 Then
    Text1.SetFocus
End If

End Sub


Private Sub Text6_Change()

If Len(Text6.Text) = 4 And Val(Text6.Text) Mod 3 = 1 Then

Image1.Visible = True
Image2.Visible = False
Image3.Visible = False
Text4.SetFocus

End If
If Len(Text6.Text) = 4 And Val(Text6.Text) Mod 3 = 0 Then
Image2.Visible = True
Image1.Visible = False
Image3.Visible = False
Text4.SetFocus

End If
If Len(Text6.Text) = 4 And Val(Text6.Text) Mod 3 = 2 Then
Image3.Visible = True
Image1.Visible = False
Image2.Visible = False
Text4.SetFocus
End If
End Sub

Private Sub Text8_Change()

If Cap1.Visible = True And Text8.Text = "W68HP" Then
Command1.Enabled = True
Text8.Locked = True
Exit Sub
Else
If Cap1.Visible = True And Len(Text8.Text) = 5 Then
MsgBox ("Captcha Incorrect.")
Text8.Text = ""
Cap2.Visible = True
Cap1.Visible = False
End If
End If

If Cap2.Visible = True And Text8.Text = "iMKiZ" Then
Command1.Enabled = True
Text8.Locked = True
Exit Sub
Else
If Cap2.Visible = True And Len(Text8.Text) = 5 Then
MsgBox ("Captcha Incorrect.")
Text8.Text = ""
Cap3.Visible = True
Cap2.Visible = False
End If
End If

If Cap3.Visible = True And Text8.Text = "qGphJD" And Len(Text8.Text) = 6 Then
Command1.Enabled = True
Text8.Locked = True
Exit Sub
Else
If Cap3.Visible = True And Len(Text8.Text) = 6 Then
MsgBox ("Captcha Incorrect.")
Text8.Text = ""
Cap4.Visible = True
Cap3.Visible = False
End If
End If

If Cap4.Visible = True And Text8.Text = "CAPTCHA" Then
Command1.Enabled = True
Text8.Locked = True
Exit Sub
Else
If Cap4.Visible = True And Len(Text8.Text) = 7 Then
    MsgBox ("Sorry. Too many wrong attempts. Application will terminate for security reasons. Login again. ")
End
End If
End If
End Sub
