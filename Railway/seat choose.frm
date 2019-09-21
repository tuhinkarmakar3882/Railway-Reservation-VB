VERSION 5.00
Begin VB.Form RESET 
   BackColor       =   &H00808000&
   Caption         =   "seat"
   ClientHeight    =   8475
   ClientLeft      =   2325
   ClientTop       =   1455
   ClientWidth     =   14160
   LinkTopic       =   "Form3"
   ScaleHeight     =   8475
   ScaleWidth      =   14160
   Begin VB.CommandButton Accept 
      Caption         =   "Accept and Continue"
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
      Left            =   6240
      TabIndex        =   37
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7700
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command26 
      Caption         =   "27"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10100
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5750
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command25 
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4300
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command23 
      Caption         =   "26"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5750
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command22 
      Caption         =   "28"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7700
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command16 
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command15 
      Caption         =   "31"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command13 
      Caption         =   "19"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10100
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command12 
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command10 
      Caption         =   "23"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command8 
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7700
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4300
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command7 
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10100
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4300
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command6 
      Caption         =   "29"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command4 
      Caption         =   "21"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command3 
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   7700
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "30"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10100
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command39 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command37 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   250
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command34 
      Caption         =   "13"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   1400
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command30 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command29 
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command24 
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   250
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command21 
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4400
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5750
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command20 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command19 
      Caption         =   "11"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5750
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command18 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   1400
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command14 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4300
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command11 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command9 
      Caption         =   "16"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7100
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   6015
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   2655
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   6135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   2655
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   6015
   End
   Begin VB.Label Label40 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SU"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   72
      Top             =   6800
      Width           =   1500
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10100
      TabIndex        =   71
      Top             =   5500
      Width           =   900
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7695
      TabIndex        =   70
      Top             =   6840
      Width           =   900
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8900
      TabIndex        =   69
      Top             =   6800
      Width           =   900
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10100
      TabIndex        =   68
      Top             =   6800
      Width           =   900
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8900
      TabIndex        =   67
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10100
      TabIndex        =   66
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SU"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   65
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8900
      TabIndex        =   64
      Top             =   5500
      Width           =   900
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8900
      TabIndex        =   63
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10100
      TabIndex        =   62
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7700
      TabIndex        =   61
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1400
      TabIndex        =   60
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2550
      TabIndex        =   59
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   58
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1400
      TabIndex        =   57
      Top             =   5505
      Width           =   900
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2550
      TabIndex        =   56
      Top             =   5500
      Width           =   900
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   250
      TabIndex        =   55
      Top             =   6800
      Width           =   900
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1400
      TabIndex        =   54
      Top             =   6800
      Width           =   900
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2550
      TabIndex        =   53
      Top             =   6800
      Width           =   900
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SU"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4400
      TabIndex        =   52
      Top             =   6800
      Width           =   1500
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SU"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4400
      TabIndex        =   51
      Top             =   4080
      Width           =   1500
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1400
      TabIndex        =   50
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   250
      TabIndex        =   49
      Top             =   4080
      Width           =   900
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   250
      TabIndex        =   48
      Top             =   5500
      Width           =   900
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4400
      TabIndex        =   47
      Top             =   5500
      Width           =   1500
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7700
      TabIndex        =   46
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7700
      TabIndex        =   45
      Top             =   5500
      Width           =   900
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4400
      TabIndex        =   44
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   43
      Top             =   5500
      Width           =   1500
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12000
      TabIndex        =   42
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   250
      TabIndex        =   41
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "seats for Senior Citizens and "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   7680
      TabIndex        =   40
      Top             =   840
      Width           =   2820
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6600
      TabIndex        =   39
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "seats for minors. Select the lower berth preferences, if available and applicable."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5100
      TabIndex        =   38
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You have applied for"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4320
      TabIndex        =   36
      Top             =   840
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   10800
      TabIndex        =   35
      Top             =   840
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   17
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR SEATS"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   14025
   End
End
Attribute VB_Name = "RESET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counts As Integer
Dim ad As Integer
Dim ch As Integer
Dim sc As Integer
Dim totalad As Integer
Dim totalval As Integer
Dim tik As Integer



Private Sub Command1_Click()
Command1.Enabled = False
Command1.BackColor = &H40&


s = Label9.Caption
Sone = Command1.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
counts = counts + 1
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If

End Sub

Private Sub Command10_Click()
Command10.Enabled = False
Command10.BackColor = &H40&
s = Label12.Caption
Sone = Command10.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
counts = counts + 1
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If

End Sub

Private Sub Command11_Click()
Command11.Enabled = False
Command11.BackColor = &H40&
counts = counts + 1
s = Label28.Caption
Sone = Command11.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If

End Sub

Private Sub Command12_Click()
Command12.Enabled = False
Command12.BackColor = &H40&
counts = counts + 1
s = Label19.Caption
Sone = Command12.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If

End Sub

Private Sub Command13_Click()
Command13.Enabled = False
Command13.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command13.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command14_Click()
Command14.Enabled = False
Command14.BackColor = &H40&
counts = counts + 1
s = Label9.Caption
Sone = Command14.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command15_Click()
Command15.Enabled = False
Command15.BackColor = &H40&
counts = counts + 1
s = Label11.Caption
Sone = Command15.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command16_Click()
Command16.Enabled = False
Command16.BackColor = &H40&
counts = counts + 1
s = Label12.Caption
Sone = Command16.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command17_Click()
Command17.Enabled = False
Command17.BackColor = &H40&
counts = counts + 1
s = Label28.Caption
Sone = Command17.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command18_Click()
Command18.Enabled = False
Command18.BackColor = &H40&
counts = counts + 1
s = Label12.Caption
Sone = Command18.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command19_Click()
Command19.Enabled = False
Command19.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command19.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
Command2.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command2.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command20_Click()
Command20.Enabled = False
Command20.BackColor = &H40&
counts = counts + 1
s = Label19.Caption
Sone = Command20.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command21_Click()
Command21.Enabled = False
Command21.BackColor = &H40&
counts = counts + 1
s = Label12.Caption
Sone = Command21.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command22_Click()
Command22.Enabled = False
Command22.BackColor = &H40&
counts = counts + 1
s = Label38.Caption
Sone = Command2.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command23_Click()
Command23.Enabled = False
Command23.BackColor = &H40&
counts = counts + 1
s = Label9.Caption
Sone = Command23.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command24_Click()
Command24.Enabled = False
Command24.BackColor = &H40&
counts = counts + 1
s = Label9.Caption
Sone = Command24.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command25_Click()
Command25.Enabled = False
Command25.BackColor = &H40&
counts = counts + 1
s = Label19.Caption
Sone = Command25.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command26_Click()
Command26.Enabled = False
Command26.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command26.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub
Private Sub Command27_Click()
Command27.Enabled = False
Command27.BackColor = &H40&
counts = counts + 1
s = Label9.Caption
Sone = Command27.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command29_Click()
Command29.Enabled = False
Command29.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command29.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
Command3.BackColor = &H40&
counts = counts + 1
s = Label9.Caption
Sone = Command3.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command30_Click()
Command30.Enabled = False
Command30.BackColor = &H40&
counts = counts + 1
s = Label28.Caption
Sone = Command30.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command34_Click()
Command34.Enabled = False
Command34.BackColor = &H40&
counts = counts + 1
s = Label28.Caption
Sone = Command34.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command37_Click()
Command37.Enabled = False
Command37.BackColor = &H40&
counts = counts + 1
s = Label9.Caption
Sone = Command37.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command39_Click()
Command39.Enabled = False
Command39.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command39.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command4_Click()
Command4.Enabled = False
Command4.BackColor = &H40&
counts = counts + 1
s = Label28.Caption
Sone = Command4.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command5_Click()
Command5.Enabled = False
Command5.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command5.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command6_Click()
Command6.Enabled = False
Command6.BackColor = &H40&
counts = counts + 1

s = Label37.Caption
Sone = Command6.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command7_Click()
Command7.Enabled = False
Command7.BackColor = &H40&
counts = counts + 1
s = Label27.Caption
Sone = Command7.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command8_Click()
Command8.Enabled = False
Command8.BackColor = &H40&
counts = counts + 1
s = Label9.Caption
Sone = Command8.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Command9_Click()
Command9.Enabled = False
Command9.BackColor = &H40&
counts = counts + 1
s = Label19.Caption
Sone = Command9.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
If counts = totalval Then
MsgBox ("Redirecting you to the Fare details Page")
Me.Hide
Form3.Show
End If
End Sub

Private Sub Form_Load()
Label2.Caption = Form2.Text18.Text
Label3.Caption = Form2.Combo7.Text
Label7.Caption = Form3.Text3.Text
Label4.Caption = Form3.Text2.Text
MDIForm1.Hide
'Form3.Show
ad = Val(Form3.Text1.Text)
ch = Val(Form3.Text2.Text)
sc = Val(Form3.Text3.Text)
counts = 0
totalad = ad
totalval = ch + sc

If totalad = 1 Then
Command5.Enabled = False
Command5.BackColor = &H40&
s = Label27.Caption
Sone = Command5.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "


End If
If totalad = 2 Then
Command17.Enabled = False
Command29.Enabled = False
Command17.BackColor = &H40&
Command29.BackColor = &H40&

s = Label25.Caption
Sone = Command17.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label21.Caption
Sone = Command29.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
End If


If totalad = 3 Then
Command13.Enabled = False
Command39.Enabled = False
Command11.Enabled = False
Command13.BackColor = &H40&
Command39.BackColor = &H40&
Command11.BackColor = &H40&

s = Label30.Caption
Sone = Command13.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label26.Caption
Sone = Command39.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label28.Caption
Sone = Command11.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "
Form3.Show
End If

If totalad = 4 Then
Command17.Enabled = False
Command34.Enabled = False
Command19.Enabled = False
Command29.Enabled = False
Command17.BackColor = &H40&
Command34.BackColor = &H40&
Command19.BackColor = &H40&
Command29.BackColor = &H40&

s = Label25.Caption
Sone = Command17.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label22.Caption
Sone = Command34.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label24.Caption
Sone = Command19.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label21.Caption
Sone = Command29.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

Form3.Show
End If
If totalad = 5 Then
Command16.Enabled = False
Command4.Enabled = False
Command23.Enabled = False
Command5.Enabled = False
Command39.Enabled = False
Command16.BackColor = &H40&
Command4.BackColor = &H40&
Command23.BackColor = &H40&
Command5.BackColor = &H40&
Command39.BackColor = &H40&

s = Label31.Caption
Sone = Command16.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label35.Caption
Sone = Command14.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label27.Caption
Sone = Command5.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label26.Caption
Sone = Command39.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label32.Caption
Sone = Command23.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "




Form3.Show
End If
If totalad = 6 Then
Command39.Enabled = False
Command20.Enabled = False
Command9.Enabled = False
Command11.Enabled = False
Command30.Enabled = False
Command17.Enabled = False
Command39.BackColor = &H40&
Command20.BackColor = &H40&
Command9.BackColor = &H40&
Command11.BackColor = &H40&
Command30.BackColor = &H40&
Command17.BackColor = &H40&

s = Label26.Caption
Sone = Command39.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label20.Caption
Sone = Command9.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label19.Caption
Sone = Command20.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label28.Caption
Sone = Command11.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label28.Caption
Sone = Command30.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

s = Label28.Caption
Sone = Command17.Caption
Form16.Label13.Caption = Form16.Label13.Caption + " " + Sone + s + ",  "

End If
'End If
'If totalval = 0 Then
'Form3.Show
'End If
End Sub





