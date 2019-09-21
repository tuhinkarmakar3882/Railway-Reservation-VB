VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Signup1 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form2"
   ClientHeight    =   10215
   ClientLeft      =   1170
   ClientTop       =   675
   ClientWidth     =   14835
   LinkTopic       =   "Form2"
   ScaleHeight     =   10215
   ScaleWidth      =   14835
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
      Height          =   615
      Left            =   3240
      TabIndex        =   19
      Top             =   9000
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc adder 
      Height          =   375
      Left            =   10800
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
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
      Caption         =   "ADDER"
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
      DataSource      =   "adder"
      Height          =   615
      Left            =   10560
      TabIndex        =   17
      Top             =   6480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSAdodcLib.Adodc signup 
      Height          =   495
      Left            =   10320
      Top             =   7200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "FINDER"
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
   Begin VB.CommandButton next1 
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
      Height          =   615
      Left            =   6240
      TabIndex        =   16
      Top             =   9000
      Width           =   2295
   End
   Begin VB.CommandButton availability 
      Caption         =   "Check Availability"
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
      Left            =   5880
      TabIndex        =   15
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Ques"
      DataSource      =   "adder"
      Height          =   315
      ItemData        =   "Signup1.frx":0000
      Left            =   3120
      List            =   "Signup1.frx":0010
      TabIndex        =   12
      Text            =   "Select Your Security Question"
      Top             =   4440
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Language"
      DataSource      =   "adder"
      Height          =   315
      ItemData        =   "Signup1.frx":008E
      Left            =   3000
      List            =   "Signup1.frx":009E
      TabIndex        =   11
      Text            =   "Select Your Language"
      Top             =   7560
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      DataField       =   "Ans"
      DataSource      =   "adder"
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
      Left            =   3120
      TabIndex        =   10
      Top             =   5640
      Width           =   1455
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
      IMEMode         =   3  'DISABLE
      Left            =   3120
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
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
      IMEMode         =   3  'DISABLE
      Left            =   3120
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      DataField       =   "userid"
      DataSource      =   "adder"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sign Up"
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
      Left            =   5520
      TabIndex        =   18
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum 8 and maximum 15 charecters Password must contain atleast one small or one capital alphabet and numeric digit"
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
      Left            =   5760
      TabIndex        =   14
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Between 3-35 charecters only letter,number and underscore are allowed"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label lang_preffered 
      BackStyle       =   0  'Transparent
      Caption         =   "Prefered Language *"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   7440
      Width           =   1695
   End
   Begin VB.Label security_ans 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Answer *"
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
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Label security_ques 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Question *"
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
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label confirm_pswrd 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password *"
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
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password *"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label id 
      BackStyle       =   0  'Transparent
      Caption         =   "User Id *"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label mandatory 
      BackStyle       =   0  'Transparent
      Caption         =   "*  Mandatory"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Signup1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Adodc1_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

Private Sub availability_Click()
Dim temp As String
temp = Text1.Text

signup.RecordSource = "select *from LoginDetails where userid='" + Text1.Text + "'"
    signup.Refresh
    If signup.Recordset.EOF Then
        MsgBox ("Accepted, you can proceed!")
        next1.Enabled = True
        Text2.Locked = False
        Text3.Locked = False
        Text1.Text = temp
    Else
        MsgBox ("User Id already in use. Please select another!")
        
    End If
End Sub

Private Sub Command1_Click()
Unload Me
Form14.Show

'adder.Recordset.AddNew
End Sub

Private Sub Form_Load()
adder.Recordset.AddNew
next1.Enabled = False

End Sub

Private Sub next1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox ("You should enter all the star (*) marked  fields")
Else
If (Text2.Text = Text3.Text) Then
    Text5.Text = Text2.Text

    adder.Recordset.Update
    Signup2.Show
    Me.Hide
Else
    MsgBox ("Passwords do not match!")
End If
End If
End Sub

