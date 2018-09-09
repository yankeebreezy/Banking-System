VERSION 5.00
Begin VB.Form Form14 
   Caption         =   "Form14"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form14"
   Picture         =   "Form14.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   13200
      TabIndex        =   12
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LAST"
      Height          =   495
      Left            =   11640
      TabIndex        =   11
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   9960
      TabIndex        =   10
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   8280
      TabIndex        =   9
      Top             =   7920
      Width           =   1455
   End
   Begin VB.TextBox Text6 
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
      Left            =   10440
      TabIndex        =   8
      Top             =   7080
      Width           =   3255
   End
   Begin VB.TextBox Text5 
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
      Left            =   10440
      TabIndex        =   7
      Top             =   6600
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   -360
      Picture         =   "Form14.frx":162EF
      ScaleHeight     =   1635
      ScaleWidth      =   20355
      TabIndex        =   4
      Top             =   -480
      Width           =   20415
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
      Left            =   10440
      TabIndex        =   3
      Top             =   6000
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
      Height          =   405
      Left            =   10440
      TabIndex        =   2
      Top             =   5400
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
      Height          =   375
      Left            =   10440
      TabIndex        =   1
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   600
      Picture         =   "Form14.frx":1F3E7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   7440
      TabIndex        =   6
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Holder"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Menu add 
      Caption         =   "Add"
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
   End
   Begin VB.Menu save 
      Caption         =   "Save"
   End
   Begin VB.Menu delete 
      Caption         =   "Delete"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db14 As Database
Dim rs14 As Recordset
Private Sub form_mouse(button As Integer, shift As Integer, x As Single, y As Single)
Set db14 = OpenDatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
Set rs14 = db14.OpenRecordset("select *from borrower")
End Sub

Private Sub add_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text5.Text = " "
Text6.Text = " "
rs14.AddNew
End Sub

Private Sub Command1_Click()
Form11.Show

Unload Form14
End Sub

Private Sub Command2_Click()
rs14.MoveFirst
movefields
End Sub

Private Sub Command3_Click()
rs14.MoveNext
If rs14.EOF Then
rs14.MoveLast
End If
movefields
End Sub

Private Sub Command4_Click()
rs14.MoveLast
movefields
End Sub

Private Sub Command5_Click()
rs14.MovePrevious
If rs14.BOF Then
rs14.MoveFirst
End If
movefields
End Sub

Private Sub delete_Click()
rs14.delete
rs14.MoveNext
If rs14.EOF Then
rs14.MoveLast
End If
movefields
End Sub

Private Sub edit_Click()
If rs14.EditMode = dbEditNone Then
rs14.edit
End If
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub save_Click()
rs14("cid") = Text1.Text
rs14("cname") = Text2.Text
rs14("loan") = Text3.Text
rs14("acchod") = Text5.Text
rs14("pwd") = Text6.Text
rs14.Update
End Sub

Private Sub movefields()
Text1.Text = rs14("cid")
Text2.Text = rs14("cname")
Text3.Text = rs14("loan")
Text5.Text = rs14("acchod")
Text6.Text = rs14("pwd")
End Sub

