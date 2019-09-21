VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Domdom 
   Caption         =   "Raya's Kitchen"
   ClientHeight    =   9615
   ClientLeft      =   720
   ClientTop       =   1065
   ClientWidth     =   19035
   LinkTopic       =   "Form5"
   ScaleHeight     =   9615
   ScaleWidth      =   19035
   Begin VB.Timer Timer1 
      Interval        =   17000
      Left            =   1320
      Top             =   9000
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20295
      ExtentX         =   35798
      ExtentY         =   19288
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
      Location        =   ""
   End
End
Attribute VB_Name = "Domdom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If Form5.Combo1.Text = "KFC" Then
WebBrowser1.Navigate ("https://online.kfc.co.in/")
Else
WebBrowser1.Navigate ("https://pizzaonline.dominos.co.in/")
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
MsgBox ("Thank You for availing our services. Do Visit Us Again.")
End
End Sub

