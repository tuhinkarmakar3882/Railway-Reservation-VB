VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fares"
   ClientHeight    =   3825
   ClientLeft      =   7275
   ClientTop       =   3375
   ClientWidth     =   5835
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5835
   Begin VB.CommandButton Command3 
      Caption         =   "Check Your Fare"
      Height          =   795
      Left            =   360
      TabIndex        =   25
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox Text17 
      Height          =   495
      Left            =   8520
      TabIndex        =   24
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text16 
      Height          =   375
      Left            =   7920
      TabIndex        =   22
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pay Now"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Confirm"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "0"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   4215
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H80000007&
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1080
            TabIndex        =   20
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label13 
            Caption         =   "Train Name"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Adult"
            Height          =   255
            Left            =   1200
            TabIndex        =   8
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H80000008&
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   1080
            TabIndex        =   7
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H80000008&
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   2160
            TabIndex        =   6
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H80000008&
            ForeColor       =   &H0000FFFF&
            Height          =   255
            Left            =   3240
            TabIndex        =   5
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Child"
            Height          =   255
            Left            =   2400
            TabIndex        =   4
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "SC"
            Height          =   255
            Left            =   3360
            TabIndex        =   3
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Fares"
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   2880
         TabIndex        =   17
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label11 
         Caption         =   "Total Fare"
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Adults"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Child"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "S.C"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.Label Label16 
      Caption         =   "class"
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "Train No."
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command2.Enabled = True
Command1.Enabled = False
'Fhandler.RecordSource = "Select Adult from fares" 'Train_No='" + Text16.Text + "' and Class='" + Text17.Text + "'"
Label12.Caption = (Val(Label5.Caption) * Val(Text1.Text)) + (Val(Label6.Caption) * Val(Text2.Text)) + (Val(Label7.Caption) * Val(Text3.Text))
End Sub

Private Sub Command2_Click()
Form19.Label2.Caption = Label12.Caption

'Temp3 = Label12.Caption
'Form5.Label11.Caption = Text1.Text
'Form5.Label13.Caption = Text2.Text
'Form5.Label15.Caption = Text3.Text
MDIForm1.Hide
Form1.Show
Me.Hide

End Sub



Private Sub Command3_Click()
Dim AdultF As Integer
Dim ChildF As Integer
Dim SeniorF As Integer
Command1.Enabled = True
Command3.Enabled = False
AdultF = 0
SeniorF = 0
ChildF = 0

If Val(Text16.Text) = 1 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 1 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 1 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 1 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 1 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 1 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 1 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 1 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 
 
 
 
If Val(Text16.Text) = 2 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 2 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 2 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 2 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 2 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 2 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 2 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 2 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 
 


If Val(Text16.Text) = 3 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 3 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 3 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 3 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 3 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 3 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 3 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 3 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 4 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 4 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 4 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 4 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 4 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 4 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 4 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 4 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If

If Val(Text16.Text) = 5 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 5 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 5 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 5 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 5 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 5 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 5 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 5 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If

If Val(Text16.Text) = 6 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 6 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 6 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 6 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 6 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 6 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 6 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 6 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If


If Val(Text16.Text) = 7 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 7 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 7 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 7 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 7 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 7 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 7 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 7 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If


If Val(Text16.Text) = 8 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 8 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 8 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 8 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 8 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 8 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 8 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 8 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If


If Val(Text16.Text) = 9 And Text17.Text = "General" Then

    AdultF = AdultF + 200
    ChildF = ChildF + 50
    SeniorF = SeniorF + 100
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 9 And Text17.Text = "I class" Then
    AdultF = AdultF + 560
    ChildF = ChildF + 300
    SeniorF = SeniorF + 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
  If Val(Text16.Text) = 9 And Text17.Text = "II class" Then
    AdultF = 500
    ChildF = 270
    SeniorF = 350
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 9 And Text17.Text = "II sitting" Then
    AdultF = 220
    ChildF = 70
    SeniorF = 120
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 9 And Text17.Text = "II sleeper" Then
    AdultF = 300
    ChildF = 150
    SeniorF = 250
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
 If Val(Text16.Text) = 9 And Text17.Text = "II Tier AC" Then
    AdultF = 560
    ChildF = 300
    SeniorF = 400
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 9 And Text17.Text = "III tier AC" Then
    AdultF = 1500
    ChildF = 750
    SeniorF = 900
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))

 End If
If Val(Text16.Text) = 9 And Text17.Text = "I AC" Then
    AdultF = 2700
    ChildF = 1700
    SeniorF = 1800
    Label5.Caption = Val(AdultF) * (Val(Text1.Text))
Label6.Caption = Val(ChildF) * (Val(Text2.Text))
Label7.Caption = Val(SeniorF) * (Val(Text3.Text))
 End If
End Sub

