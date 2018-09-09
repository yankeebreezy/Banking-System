VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   600
      Picture         =   "Form5.frx":16338
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11040
      TabIndex        =   4
      Top             =   6360
      Width           =   3255
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
      Height          =   420
      Left            =   11040
      TabIndex        =   3
      Top             =   5895
      Width           =   3255
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
      Height          =   420
      Left            =   11040
      TabIndex        =   2
      Top             =   5400
      Width           =   3255
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
      Height          =   420
      Left            =   11040
      TabIndex        =   1
      Top             =   4920
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form5.frx":19A62
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
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db5 As Database
Dim rs5 As Recordset
Dim a As String
Private Sub Form_Load()
Set db5 = opendatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
a = Form3.Text2
Set rs5 = db5.openrecordset("select cid,cname,street,city from customer where pwd=" & a)
Text1.Text = rs5.fields("cid")
Text2.Text = rs5.fields("cname")
Text3.Text = rs5.fields("street")
Text4.Text = rs5.fields("city")
End Sub
Private Sub Command1_Click()
Form4.Show

Unload Form5
End Sub

Private Sub exit_Click()
End
End Sub


