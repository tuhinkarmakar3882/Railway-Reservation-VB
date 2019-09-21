VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Train"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   3200
      Width           =   1695
   End
   Begin VB.ComboBox Combo7 
      BackColor       =   &H000080FF&
      Height          =   315
      ItemData        =   "Form4.frx":0000
      Left            =   5760
      List            =   "Form4.frx":0019
      TabIndex        =   14
      Text            =   "Select Destination"
      Top             =   700
      Width           =   1935
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H000080FF&
      Height          =   315
      ItemData        =   "Form4.frx":0064
      Left            =   1800
      List            =   "Form4.frx":0080
      TabIndex        =   13
      Text            =   "Select Source"
      Top             =   700
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3200
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8040
      Top             =   3480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "Trains"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Book Ticket"
      Height          =   2535
      Left            =   6120
      TabIndex        =   11
      Top             =   1275
      Width           =   1455
   End
   Begin VB.ComboBox Combo5 
      DataField       =   "OP_Basis"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   13200
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      DataField       =   "To"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   12960
      TabIndex        =   4
      Top             =   5040
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "From"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   11400
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Train_Name"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   9120
      TabIndex        =   2
      Top             =   5040
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form4.frx":00CF
      Left            =   13320
      List            =   "Form4.frx":00EE
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Train No."
      Height          =   495
      Left            =   150
      TabIndex        =   21
      Top             =   1400
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Train Name:"
      Height          =   495
      Left            =   150
      TabIndex        =   20
      Top             =   2200
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Destination :"
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   700
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Source :"
      Height          =   375
      Left            =   150
      TabIndex        =   18
      Top             =   700
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1800
      TabIndex        =   16
      Top             =   2205
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   1800
      TabIndex        =   15
      Top             =   1395
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "OP_Basis"
      Height          =   375
      Left            =   12000
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   13080
      TabIndex        =   9
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "From"
      Height          =   255
      Left            =   11520
      TabIndex        =   8
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Train Name"
      Height          =   255
      Left            =   9120
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Enter Train Number"
      Height          =   255
      Left            =   11280
      TabIndex        =   6
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Train"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   20
      Width           =   7815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Integer


Private Sub Combo1_Click()
Adodc1.Refresh
Adodc1.Recordset.Find "Train_No =" & Combo1.Text, 0, adSearchForward
If Adodc1.Recordset.EOF = True Then
MsgBox ("Train not Available")
End If
End Sub


Private Sub Combo2_Change()
'Temp2 = Combo2.Text
'Form2.Text18 = Temp2
'Form2.Show
End Sub

Private Sub Combo6_LostFocus()
Command1.Enabled = False

End Sub

Private Sub Combo7_LostFocus()
Command1.Enabled = False
End Sub

Private Sub Command1_Click()

Temp1 = Label7.Caption
Form2.Text1 = Temp1
Temp2 = Label8.Caption
Form2.Text18 = Temp2
temp3 = Combo6.Text
Form2.Text19 = temp3
Temp4 = Combo7.Text
Form2.Text20 = Temp4


Form2.Show
Me.Hide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
flag = 0
If Combo6.Text = "HOWRAH" And Combo7.Text = "AMRITSAR" Then
flag = 1
Label7.Caption = "1"
Label8.Caption = "HOWRAH-AMRTISAR EXPRESS"
End If
If Combo6.Text = "MALDA" And Combo7.Text = "BHIWANI" Then
flag = 1
Label7.Caption = "2"
Label8.Caption = "MALDA-BHIWANI EXPRESS"
End If
If Combo6.Text = "ALLAHABAD" And Combo7.Text = "FAIZABAD" Then
flag = 1
Label7.Caption = "3"
Label8.Caption = "ALLAHABAD-FAIZABAD EXPRESS"
End If
If Combo6.Text = "HOWRAH" And Combo7.Text = "JAMMUTAWAI" Then
flag = 1
Label7.Caption = "4"
Label8.Caption = "HOWRAH-JAMMUTAWAI EXPRESS"
End If
If Combo6.Text = "VARANASI" And Combo7.Text = "LUCKNOW" Then
flag = 1
Label7.Caption = "5"
Label8.Caption = "VARANASI-LUCKNOW EXPRESS"
End If
If Combo6.Text = "BOMBAY" And Combo7.Text = "FAIZABAD" Then
flag = 1
Label7.Caption = "6"
Label8.Caption = "BOMBAY-FAIZABAD-SAKET EXPRESS"
End If
If Combo6.Text = "PATNA" And Combo7.Text = "NEW DELHI" Then
flag = 1
Label7.Caption = "7"
Label8.Caption = "PATNA-NEW DELHI EXPRESS"
End If
If Combo6.Text = "MUZAFFARPUR" And Combo7.Text = "NEW DELHI" Then
flag = 1
Label7.Caption = "8"
Label8.Caption = "MUZAFFARPUR-NEW DELHI EXPRESS"
End If
If Combo6.Text = "SULTANPUR" And Combo7.Text = "NEW DELHI" Then
flag = 1
Label7.Caption = "9"
Label8.Caption = "SULTANPUR-NEW DELHI EXPRESS"
End If

If Combo6.Text = "ALLAHABAD" And Combo7.Text = "LUCKNOW" Then
flag = 1
Label7.Caption = "5"
Label8.Caption = "VARANASI-LUCKNOW EXPRESS"
End If

If flag = 0 And (Combo6.Text = "Select Source" Or Combo7.Text = "Select Destination") Then
MsgBox ("Error! Make sure you have selected source and destination properly.")
Command1.Enabled = False
Else
If flag = 0 And Combo6.Text <> "Select Source" And Combo7.Text <> "Select Destination" Then
'MsgBox ("Alert! There is currently no direct train running on this route. But your tickets will be booked for HOWRAH-JAMMUTAWAI EXPRESS via your destination.")
Label7.Caption = "4"
Label8.Caption = "HOWRAH-JAMMUTAWAI EXPRESS"
Command1.Enabled = True
Else
flag = 0
Command1.Enabled = True
End If
End If

End Sub

Private Sub Form_Load()
Form2.Hide
Command1.Enabled = False
End Sub

