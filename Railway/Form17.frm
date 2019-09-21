VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   7350
   ClientLeft      =   4140
   ClientTop       =   2655
   ClientWidth     =   7860
   LinkTopic       =   "Form17"
   Picture         =   "Form17.frx":0000
   ScaleHeight     =   7350
   ScaleWidth      =   7860
   Begin MSAdodcLib.Adodc PHandler 
      Height          =   495
      Left            =   5280
      Top             =   5280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
   Begin VB.ComboBox Combo2 
      DataField       =   "Ques"
      DataSource      =   "adder"
      Height          =   315
      ItemData        =   "Form17.frx":BAC5
      Left            =   3200
      List            =   "Form17.frx":BAD5
      TabIndex        =   7
      Text            =   "Select Your Security Question"
      Top             =   2280
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Validate"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3200
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3200
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD RECOVERY"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Ans :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1000
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Question :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1005
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Id :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1000
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PHandler.RecordSource = "select *from LoginDetails where userid='" + Text1.Text + "' and Ques='" + Combo2.Text + "' and Ans ='" + Text2.Text + "'"
PHandler.Refresh
If PHandler.Recordset.EOF Then
        MsgBox ("Wrong Credentials")
Else
        Form18.Show
        Form18.Text4.Text = Text1.Text
        Unload Me
End If
End Sub

