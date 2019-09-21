VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Railway Reservation System"
   ClientHeight    =   7335
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7395
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5062.747
   ScaleMode       =   0  'User
   ScaleWidth      =   6944.287
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser Wb1 
      Height          =   5895
      Left            =   600
      TabIndex        =   4
      Top             =   2640
      Width           =   7335
      ExtentX         =   12938
      ExtentY         =   10398
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   2640
      Left            =   3840
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "'2-R-S-T' Railway Services"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1822.175
      Y2              =   1822.175
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":171EC
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1650
      Left            =   11
      TabIndex        =   1
      Top             =   480
      Width           =   3765
   End
   Begin VB.Label lblTitle 
      Caption         =   "'2-R-S-T' Railway Services"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   480
      TabIndex        =   2
      Top             =   -960
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1822.175
      Y2              =   1822.175
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()

Form14.Show

Unload Me
End Sub



Private Sub Form_Load()
Wb1.Navigate ("E:\Railway\Images\ani.gif")
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

End Sub

