VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form16"
   Picture         =   "Form16.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   13320
      TabIndex        =   12
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LAST"
      Height          =   495
      Left            =   11640
      TabIndex        =   11
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   10080
      TabIndex        =   10
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   8280
      TabIndex        =   9
      Top             =   7680
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
      Left            =   10560
      TabIndex        =   8
      Top             =   6960
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
      Left            =   10560
      TabIndex        =   7
      Top             =   6480
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form16.frx":16258
      ScaleHeight     =   1635
      ScaleWidth      =   20235
      TabIndex        =   4
      Top             =   -360
      Width           =   20295
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
      Left            =   10560
      TabIndex        =   3
      Top             =   4800
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
      Height          =   375
      Left            =   10560
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
      Left            =   10560
      TabIndex        =   1
      Top             =   6000
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   600
      Picture         =   "Form16.frx":1F350
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1680
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
      Top             =   6960
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
      Top             =   6480
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
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db16 As Database
Dim rs16 As Recordset
Private Sub form_mouse(button As Integer, shift As Integer, x As Single, y As Single)
Set db16 = OpenDatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
Set rs16 = db16.OpenRecordset("select *from depositer")
End Sub

Private Sub add_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text5.Text = " "
Text6.Text = " "
rs16.AddNew
End Sub

Private Sub Command1_Click()
Form11.Show

Unload Form16
End Sub

Private Sub Command2_Click()
rs16.MoveFirst
movefields
End Sub

Private Sub Command3_Click()
rs16.MoveNext
If rs16.EOF Then
rs16.MoveLast
End If
movefields
End Sub

Private Sub Command4_Click()
rs16.MoveLast
movefields
End Sub

Private Sub Command5_Click()
rs16.previous
If rs16.BOF Then
rs16.MoveFirst
End If
movefields
End Sub

Private Sub delete_Click()
rs16.delete
rs16.MoveNext
If rs16.EOF Then
rs16.MoveLast
End If
movefields
End Sub

Private Sub edit_Click()
If rs16.EditMode = dbEditNone Then
rs16.eidt
End If
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub save_Click()
rs16("cid") = Text1.Text
rs16("cname") = Text2.Text
rs16("accno") = Text3.Text
rs16("acchod") = Text5.Text
rs16("pwd") = Text6.Text
rs16.Update
End Sub

Private Sub movefields()
Text1.Text = rs16("cid")
Text2.Text = rs16("cname")
Text3.Text = rs16("accno")
Text5.Text = rs16("acchod")
Text6.Text = rs16("pwd")
End Sub
