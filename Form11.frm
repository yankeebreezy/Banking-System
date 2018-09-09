VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   10650
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      Picture         =   "Form11.frx":5221F
      ScaleHeight     =   1515
      ScaleWidth      =   20235
      TabIndex        =   2
      Top             =   -480
      Width           =   20295
   End
   Begin VB.PictureBox Picture2 
      Height          =   780
      Left            =   720
      Picture         =   "Form11.frx":5B317
      ScaleHeight     =   720
      ScaleMode       =   0  'User
      ScaleWidth      =   1041.509
      TabIndex        =   1
      Top             =   1440
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   1800
      Picture         =   "Form11.frx":5EF2F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USER ACCESS INFO"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   9000
      TabIndex        =   9
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER DETAILS"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   7200
      TabIndex        =   8
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEPOSITER DETAILS"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   855
      Left            =   10080
      TabIndex        =   7
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOAN DETAILS"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   12720
      TabIndex        =   6
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BORROWER DETAILS"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   975
      Left            =   10200
      TabIndex        =   5
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BRANCH DETAILS"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1095
      Left            =   7080
      TabIndex        =   4
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT DETAILS"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1335
      Left            =   4680
      TabIndex        =   3
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show

Unload Form11
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Label1_Click()
Form12.Show

Unload Form11
End Sub

Private Sub Label2_Click()
Form16.Show

Unload Form11
End Sub

Private Sub Label3_Click()
Form17.Show

Unload Form11
End Sub

Private Sub Label4_Click()
Form14.Show

Unload Form11
End Sub

Private Sub Label5_Click()
Form15.Show

Unload Form11
End Sub

Private Sub Label6_Click()
Form13.Show

Unload Form11
End Sub

Private Sub Label7_Click()
Form18.Show

Unload Form11
End Sub

Private Sub Picture2_Click()
Form1.Show

Unload Form11
End Sub
