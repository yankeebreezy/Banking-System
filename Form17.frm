VERSION 5.00
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form17"
   Picture         =   "Form17.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   8280
      TabIndex        =   12
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   12960
      TabIndex        =   11
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LAST"
      Height          =   495
      Left            =   11400
      TabIndex        =   10
      Top             =   7920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   7920
      Width           =   1335
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
      Left            =   10320
      TabIndex        =   8
      Top             =   7200
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
      Left            =   10320
      TabIndex        =   7
      Top             =   6720
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   -120
      Picture         =   "Form17.frx":15CD3
      ScaleHeight     =   1635
      ScaleWidth      =   20475
      TabIndex        =   4
      Top             =   -480
      Width           =   20535
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   10320
      TabIndex        =   3
      Top             =   4680
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   10320
      TabIndex        =   2
      Top             =   5400
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   6120
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   600
      Picture         =   "Form17.frx":1EDCB
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
      Left            =   7320
      TabIndex        =   6
      Top             =   7200
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
      Left            =   7320
      TabIndex        =   5
      Top             =   6720
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
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db17 As Database
Dim rs17 As Recordset
Private Sub form_mouse(button As Integer, shift As Integer, x As Single, y As Single)
Set db17 = OpenDatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
Set rs17 = db17.OpenRecordset("select *from loan")
End Sub
Private Sub add_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text5.Text = " "
Text6.Text = " "
rs17.AddNew
End Sub

Private Sub Command1_Click()
Form11.Show

Unload Form17
End Sub

Private Sub Command2_Click()
rs17.MoveNext
If rs17.EOF Then
rs17.MoveLast
End If
movefields
End Sub

Private Sub Command3_Click()
rs17.MoveLast
movefields
End Sub

Private Sub Command4_Click()
rs17.MovePrevious
If rs17.BOF Then
rs17.MoveFirst
End If
movefields
End Sub

Private Sub Command5_Click()
rs17.MoveFirst
movefields
End Sub

Private Sub delete_Click()
rs17.delete
rs17.MoveNext
If rs.EOF Then
rs.MoveLast
End If
movefields
End Sub

Private Sub edit_Click()
If rs17.EditMode = dbEditNone Then
rs17.eidt
End If
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub save_Click()
rs17("lno") = Text1.Text
rs17("bname") = Text2.Text
rs17("acc") = Text3.Text
rs17("acchod") = Text5.Text
rs17("pwd") = Text6.Text
rs17.Update
End Sub
Private Sub movefields()
Text1.Text = rs17("lno")
Text2.Text = rs17("bname")
Text3.Text = rs17("acc")
Text5.Text = rs17("acchod")
Text6.Text = rs17("pwd")
End Sub
