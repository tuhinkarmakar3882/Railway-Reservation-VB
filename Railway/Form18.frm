VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form18 
   Caption         =   "Form18"
   ClientHeight    =   4800
   ClientLeft      =   5550
   ClientTop       =   2055
   ClientWidth     =   6810
   LinkTopic       =   "Form18"
   Picture         =   "Form18.frx":0000
   ScaleHeight     =   4800
   ScaleWidth      =   6810
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   735
      Left            =   9120
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   8760
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "LoginDetails"
      Caption         =   "Adodc1"
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
   Begin VB.TextBox Text5 
      DataField       =   "pass"
      DataSource      =   "Adodc1"
      Height          =   615
      Left            =   9720
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   9000
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "pass"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3000
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Above Password"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Type New Password"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD RESET"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6495
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Refresh
Adodc1.Recordset.Find "userid= '" + Text4.Text + "'", 0, adSearchForward
If Adodc1.Recordset.EOF = True Then
MsgBox ("Not found")
End If
If Text2.Text = Text1.Text Then
     Text3.Text = Text1.Text
     Adodc1.Recordset.AddNew
     MsgBox ("Password Reset Successfully!")
     Me.Hide
     Form14.Show
Else
    MsgBox ("Passwords Do Not Match")
End If
End Sub

Private Sub Command2_Click()
Adodc1.Refresh
Adodc1.Recordset.Find "userid= '" + Text4.Text + "'", 0, adSearchForward
If Adodc1.Recordset.EOF = True Then
MsgBox ("Not found")
End If
End Sub

