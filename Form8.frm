VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   600
      Picture         =   "Form8.frx":15D4C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   3
      Top             =   5880
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10200
      TabIndex        =   1
      Top             =   4560
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form8.frx":19476
      ScaleHeight     =   1635
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   -480
      Width           =   20295
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db8 As Database
Dim rs8 As Recordset
Dim d As String

Private Sub Command1_Click()
Form4.Show

Unload Form8
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()

Set db8 = opendatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
d = Form3.Text2
Set rs8 = db8.openrecordset("select bname,bcity,assests from branch where pwd=" & d)
Text1.Text = rs8.fields("accno")
Text2.Text = rs8.fields("bname")
Text3.Text = rs8.fields("bal")

End Sub
