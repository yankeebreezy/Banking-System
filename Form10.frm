VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "Form10.frx":15CD3
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   10080
      TabIndex        =   3
      Top             =   6120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   10080
      TabIndex        =   2
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   10080
      TabIndex        =   1
      Top             =   4680
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      Picture         =   "Form10.frx":193FD
      ScaleHeight     =   1635
      ScaleWidth      =   20475
      TabIndex        =   0
      Top             =   -480
      Width           =   20535
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db10 As Database
Dim rs10 As Recordset
Dim f As String

Private Sub Form_Load()
Set db10 = opendatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
f = Form3.Text2
Set rsf = dbf.openrecordset("select lno,bname,acc from loan where pwd=" & f)
Text1.Text = rs.fields("lno")
Text2.Text = rs.fields("bname")
Text3.Text = rs.fields("acc")
End Sub

Private Sub Command1_Click()
Form4.Show

Unload Form10
End Sub

Private Sub exit_Click()
End
End Sub
