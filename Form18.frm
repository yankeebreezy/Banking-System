VERSION 5.00
Begin VB.Form Form18 
   Caption         =   "Form18"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form18"
   Picture         =   "Form18.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Height          =   735
      Left            =   600
      Picture         =   "Form18.frx":5221F
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   13800
      TabIndex        =   9
      Top             =   6720
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "LAST"
      Height          =   495
      Left            =   12480
      TabIndex        =   8
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   11160
      TabIndex        =   7
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   9840
      TabIndex        =   6
      Top             =   6720
      Width           =   1095
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
      TabIndex        =   4
      Top             =   5760
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
      TabIndex        =   3
      Top             =   5160
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form18.frx":55949
      ScaleHeight     =   1635
      ScaleWidth      =   20235
      TabIndex        =   0
      Top             =   -480
      Width           =   20295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESS INFO"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   735
      Left            =   9000
      TabIndex        =   5
      Top             =   4080
      Width           =   3735
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
      Left            =   7200
      TabIndex        =   2
      Top             =   5760
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
      Left            =   7200
      TabIndex        =   1
      Top             =   5040
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
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db18 As Database
Dim rs18 As Recordset
Private Sub form_mouse(button As Integer, shift As Integer, x As Single, y As Single)
Set db18 = OpenDatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
Set rs18 = db18.OpenRecordset("select *from accessinfo")
End Sub

Private Sub add_Click()
Text5.Text = " "
Text6.Text = " "
rs18.AddNew
End Sub

Private Sub Command2_Click()
rs18.MoveNext
If rs18.EOF Then
rs18.MoveLast
End If
movefields
End Sub

Private Sub Command3_Click()
rs18.MoveLast
movefields
End Sub

Private Sub Command4_Click()
rs18.MovePrevious
If rs18.BOF Then
rs18.MoveFirst
End If
movefields
End Sub

Private Sub Command5_Click()
Form11.Show

Unload Form18
End Sub

Private Sub delete_Click()
rs18.delete
rs18.MoveNext
If rs18.EOF Then
rs18.MoveLast
End If
movefields
End Sub

Private Sub edit_Click()
If rs18.EditMode = dbEditNone Then
rs18.edit
End If
End Sub


Private Sub Command1_Click()
rs18.MoveFirst
movefields
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub save_Click()
rs18("acchod") = Text5.Text
rs18("pwd") = Text6.Text
rs18.Update
End Sub

Private Sub movefields()
Text5.Text = rs18("acchod")
Text6.Text = rs18("pwd")
End Sub
