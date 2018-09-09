VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   600
      Picture         =   "Form7.frx":162EF
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
      Left            =   10320
      TabIndex        =   3
      Top             =   4800
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
      Height          =   405
      Left            =   10320
      TabIndex        =   2
      Top             =   5400
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
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   6000
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form7.frx":19A19
      ScaleHeight     =   1635
      ScaleWidth      =   20355
      TabIndex        =   0
      Top             =   -480
      Width           =   20415
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db7 As Database
Dim rs7 As Recordset
Dim c As String

Private Sub Form_Load()
Set db7 = opendatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
c = Form3.Text2
Set rs7 = db7.openrecordset("select cid,cname,loan from borrower where pwd=" & c)
Text1.Text = rs7.fields("cid")
Text2.Text = rs7.fields("cname")
Text3.Text = rs7.fields("loan")
End Sub
Private Sub Command1_Click()
Form4.Show

Unload Form7
End Sub

Private Sub exit_Click()
End
End Sub
