VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "Form9"
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
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
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
      Picture         =   "Form9.frx":16258
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   10560
      TabIndex        =   3
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10560
      TabIndex        =   1
      Top             =   4800
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
      Picture         =   "Form9.frx":19982
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
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db9 As Database
Dim rs9 As Recordset
Dim e As String

Private Sub Form_Load()
Set db9 = opendatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
e = Form3.Text2
Set rs8 = db8.openrecordset("select cid,cname,accno from depositer where pwd=" & e)
Text1.Text = rs8.fields("cid")
Text2.Text = rs8.fields("cname")
Text3.Text = rs8.fields("accno")
End Sub

Private Sub Command1_Click()
Form4.Show

Unload Form9
End Sub

Private Sub exit_Click()
End
End Sub
