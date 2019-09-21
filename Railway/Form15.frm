VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   2430
   ClientLeft      =   5940
   ClientTop       =   3840
   ClientWidth     =   4575
   LinkTopic       =   "Form15"
   ScaleHeight     =   2430
   ScaleWidth      =   4575
   Begin VB.CommandButton Command2 
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ARE YOU SURE??"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd2_Click()
Form14.Show
Unload Me
End Sub

Private Sub Command2_Click()
End
End Sub
