VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   720
      Picture         =   "Form4.frx":5221F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   1800
      Picture         =   "Form4.frx":55E37
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   0
      Picture         =   "Form4.frx":59561
      ScaleHeight     =   1515
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   -480
      Width           =   20295
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
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
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
      TabIndex        =   5
      Top             =   4800
      Width           =   2535
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
      Left            =   10080
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
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
      TabIndex        =   3
      Top             =   4800
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
      TabIndex        =   2
      Top             =   3480
      Width           =   2295
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
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Form3.Show

Unload Form4
End Sub

Private Sub Command2_Click()
Form1.Show

Unload Form4
End Sub

Private Sub exit_Click()
End
End Sub




Private Sub Label1_Click()
Form5.Show

Unload Form4
End Sub

Private Sub Label2_Click()
Form9.Show

Unload Form4
End Sub



Private Sub Label3_Click()
Form10.Show

Unload Form4
End Sub

Private Sub Label4_Click()
Form7.Show

Unload Form4
End Sub

Private Sub Label5_Click()
Form8.Show

Unload Form4
End Sub

Private Sub Label6_Click()
Form6.Show

Unload Form4
End Sub

Private Sub Picture2_Click()
Form1.Show

Unload Form4
End Sub
