VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fare Details"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4950
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Fares"
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox Combo7 
      DataSource      =   "Adodc1"
      Height          =   315
      ItemData        =   "Form13.frx":0000
      Left            =   1080
      List            =   "Form13.frx":001C
      TabIndex        =   14
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Train Name"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "SC"
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Child"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Adult"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "Train Number"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label Label11 
      Caption         =   "Enter Class"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Enter Train Number"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
If Text1.Text <> "" And Combo7.Text <> "" Then
Label14.Caption = Text1.Text
'Label10.Caption = Combo7.Text
Select Case Label14.Caption
Case "1", "2", "3", "4", "5", "6", "7", "8", "9"
Select Case Combo7.Text
Case "General"
Label5.Caption = "200"
Label6.Caption = "50"
Label7.Caption = "100"
Case "II class"
Label5.Caption = "500"
Label6.Caption = "270"
Label7.Caption = "350"
Case "II sitting"
Label5.Caption = "220"
Label6.Caption = "70"
Label7.Caption = "120"
Case "II sleeper"
Label5.Caption = "300"
Label6.Caption = "150"
Label7.Caption = "250"
Case "I class"
Label5.Caption = "560"
Label6.Caption = "300"
Label7.Caption = "400"
Case "III tier AC"
Label5.Caption = "1500"
Label6.Caption = "750"
Label7.Caption = "900"
Case "II Tier AC"
Label5.Caption = "560"
Label6.Caption = "300"
Label7.Caption = "400"
Case "I AC"
Label5.Caption = "2750"
Label6.Caption = "1700"
Label7.Caption = "1800"
'Else
'MsgBox ("Please do not leave any field blank")
'End If
End Select
End Select
End If
If Text1.Text = 1 Then
Label2.Caption = "HOWRAH-AMRITSAR EXPRESS"
End If
If Text1.Text = 2 Then
Label2.Caption = "MALDA-BHIWANI EXPRESS"
End If
If Text1.Text = 3 Then
Label2.Caption = "ALLAHABAD-FAIZABAD EXPRESS"
End If
If Text1.Text = 4 Then
Label2.Caption = "HOWRAH-JAMMUTAWAI EXPRESS"
End If
If Text1.Text = 6 Then
Label2.Caption = "BOMBAY-FAIZABAD SAKET EXPRESS"
End If
If Text1.Text = 5 Then
Label2.Caption = "VARANASI-LUCKNOW EXPRESS"
End If
If Text1.Text = 7 Then
Label2.Caption = "PATNA-NEW DELHI EXPRESS"
End If
If Text1.Text = 8 Then
Label2.Caption = "MUZAFFARPUR-NEW DELHI EXPRESS"
End If
If Text1.Text = 9 Then
Label2.Caption = "SULTANPUR-DELHI EXPRESS"
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
