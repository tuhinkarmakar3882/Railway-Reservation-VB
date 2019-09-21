VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form16 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form 16"
   ClientHeight    =   8955
   ClientLeft      =   4050
   ClientTop       =   1785
   ClientWidth     =   11940
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form16.frx":0000
   ScaleHeight     =   8955
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   720
      Top             =   240
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form16.frx":1661D
      Height          =   1695
      Left            =   12240
      TabIndex        =   45
      Top             =   6840
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text7 
      DataField       =   "dt"
      DataSource      =   "tst"
      Height          =   495
      Left            =   14760
      TabIndex        =   44
      Text            =   "dt"
      Top             =   6120
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      DataField       =   "at"
      DataSource      =   "tst"
      Height          =   495
      Left            =   14880
      TabIndex        =   43
      Text            =   "at"
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      DataField       =   "doj"
      DataSource      =   "tst"
      Height          =   375
      Left            =   15000
      TabIndex        =   42
      Text            =   "doj"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      DataField       =   "des"
      DataSource      =   "tst"
      Height          =   615
      Left            =   16440
      TabIndex        =   41
      Text            =   "des"
      Top             =   3840
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "src"
      DataSource      =   "tst"
      Height          =   615
      Left            =   16440
      TabIndex        =   40
      Text            =   "src"
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "pnr"
      DataSource      =   "tst"
      Height          =   495
      Left            =   14640
      TabIndex        =   39
      Text            =   "pnr"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      DataField       =   "user"
      DataSource      =   "tst"
      Height          =   495
      Left            =   15840
      TabIndex        =   38
      Text            =   "user"
      Top             =   1560
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc tst 
      Height          =   375
      Left            =   15480
      Top             =   720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      RecordSource    =   "select * from Table2"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   " '2R-S-T' Railway Services"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   960
      TabIndex        =   37
      Top             =   1320
      Width           =   10335
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "PNR NO :"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6150
      TabIndex        =   36
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7640
      TabIndex        =   35
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   10080
      TabIndex        =   34
      Top             =   6300
      Width           =   795
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Arr. Time :"
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   8760
      TabIndex        =   33
      Top             =   6300
      Width           =   795
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   400
      Left            =   7640
      TabIndex        =   32
      Top             =   6300
      Width           =   800
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Dep. Time :"
      ForeColor       =   &H0000FFFF&
      Height          =   400
      Left            =   6150
      TabIndex        =   31
      Top             =   6300
      Width           =   820
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Journey :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6150
      TabIndex        =   30
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label31 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label31"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7640
      TabIndex        =   29
      Top             =   5340
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Label12"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7635
      TabIndex        =   28
      Top             =   4680
      Width           =   3255
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Coach :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6150
      TabIndex        =   27
      Top             =   5340
      Width           =   1095
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3500
      TabIndex        =   26
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4500
      TabIndex        =   25
      Top             =   3000
      Width           =   600
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4500
      TabIndex        =   24
      Top             =   3800
      Width           =   600
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4500
      TabIndex        =   23
      Top             =   4600
      Width           =   600
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4500
      TabIndex        =   22
      Top             =   5400
      Width           =   600
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4500
      TabIndex        =   21
      Top             =   6200
      Width           =   600
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4500
      TabIndex        =   20
      Top             =   6950
      Width           =   600
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3500
      TabIndex        =   19
      Top             =   3800
      Width           =   600
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3500
      TabIndex        =   18
      Top             =   4600
      Width           =   600
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3500
      TabIndex        =   17
      Top             =   5400
      Width           =   600
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3500
      TabIndex        =   16
      Top             =   6200
      Width           =   600
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3500
      TabIndex        =   15
      Top             =   6950
      Width           =   600
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seats :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6150
      TabIndex        =   14
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Destination :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6120
      TabIndex        =   13
      Top             =   4700
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Source :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   6150
      TabIndex        =   12
      Top             =   4000
      Width           =   1455
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Wish You a Happy and Safe Jouney !"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   7560
      Width           =   11880
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   7635
      TabIndex        =   10
      Top             =   6840
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7640
      TabIndex        =   9
      Top             =   4005
      Width           =   3495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   6480
      TabIndex        =   7
      Top             =   2520
      Width           =   3975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   490
      Left            =   1500
      TabIndex        =   6
      Top             =   6950
      Width           =   1800
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   490
      Left            =   1500
      TabIndex        =   5
      Top             =   6200
      Width           =   1800
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   490
      Left            =   1500
      TabIndex        =   4
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   490
      Left            =   1500
      TabIndex        =   3
      Top             =   4600
      Width           =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   490
      Left            =   1500
      TabIndex        =   2
      Top             =   3800
      Width           =   1800
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FFFF&
      Height          =   490
      Left            =   1500
      TabIndex        =   1
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Passenger Details :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Form5.Show

End Sub
