VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSplash 
   Caption         =   "Login"
   ClientHeight    =   2880
   ClientLeft      =   6165
   ClientTop       =   4740
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4185
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
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "FORGOT  PASSWORD"
      Height          =   975
      Left            =   2880
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Lhandler 
      Height          =   495
      Left            =   2160
      Top             =   5520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from LoginDetails"
      Caption         =   "Login_kumar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   2160
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Username"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Image Cap2 
      Height          =   795
      Left            =   600
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image Cap4 
      Height          =   795
      Left            =   600
      Picture         =   "Form1.frx":7513
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image Cap3 
      Height          =   795
      Left            =   600
      Picture         =   "Form1.frx":CC07
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.Image Cap1 
      Height          =   795
      Left            =   600
      Picture         =   "Form1.frx":1F0A1
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1995
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Lhandler.RecordSource = "select *from LoginDetails where userid='" + Text3.Text + "' and pass='" + Text4.Text + "'"
    Lhandler.Refresh
    If Lhandler.Recordset.EOF Then
        Text8.Locked = False
        Text8.Text = ""
        If Cap1.Visible = True Then
        Cap2.Visible = False
        Cap1.Visible = False
        Cap3.Visible = True
        End If
        If Cap2.Visible = True Then
        Cap2.Visible = False
        Cap1.Visible = True
        Cap3.Visible = False
        End If
        If Cap3.Visible = True Then
        Cap2.Visible = True
        Cap1.Visible = False
        Cap3.Visible = False
        End If
        MsgBox ("Not Registered! Please Register!")
        Command1.Enabled = False
        Text4.Text = ""
    Else
        MsgBox ("You have successfully logged in!")
        'Form16.tst.Recordset.AddNew
        Form16.Text1.Text = Text3.Text
        Form20.Text1.Text = Text3.Text
        MDIForm1.Text1.Text = Text3.Text
        Form20.Cancel.RecordSource = "select *from table2 where user='" + Form20.Text1.Text + "'"
        Form20.Cancel.Refresh
        MDIForm1.Show
        Unload Me

    End If
End Sub

Private Sub Command2_Click()
Form14.Show
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
Form17.Show

End Sub

Private Sub Form_Load()
Cap1.Visible = False
Cap2.Visible = False
Cap3.Visible = True
Cap4.Visible = False
Command1.Enabled = False
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
Cap3.Visible = True
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
Cap1.Visible = True
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
Cap2.Visible = True
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
