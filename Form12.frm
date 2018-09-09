VERSION 5.00
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   13560
      TabIndex        =   13
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "LAST"
      Height          =   495
      Left            =   12000
      TabIndex        =   12
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   10440
      TabIndex        =   11
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   8640
      TabIndex        =   10
      Top             =   8520
      Width           =   1575
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
      Left            =   10920
      TabIndex        =   9
      Top             =   7680
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
      Left            =   10920
      TabIndex        =   8
      Top             =   7200
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   600
      Picture         =   "Form12.frx":16338
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "Form12.frx":19A62
      ScaleHeight     =   1635
      ScaleWidth      =   20235
      TabIndex        =   4
      Top             =   -480
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
      Height          =   420
      Left            =   10920
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
      Height          =   420
      Left            =   10920
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
      Height          =   420
      Left            =   10920
      TabIndex        =   1
      Top             =   6000
      Width           =   3255
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
      Left            =   10920
      TabIndex        =   0
      Top             =   6600
      Width           =   3255
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
      Left            =   7920
      TabIndex        =   7
      Top             =   7680
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
      Left            =   7920
      TabIndex        =   6
      Top             =   7200
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
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db12 As Database
Dim rs12 As Recordset
Private Sub form_mouse(button As Integer, shift As Integer, x As Single, y As Single)

Set db12 = OpenDatabase("banksys", False, False, "odbc;uid=cse2a1;pwd=banksys;dsn=banksys")
Set rs12 = db12.OpenRecordset("select *from customer")
End Sub




Private Sub add_Click()
Text1.Text = " "
Text2.Text = " "
Text3.Text = " "
Text4.Text = " "
Text5.Text = " "
Text6.Text = " "
rs12.AddNew
End Sub

Private Sub Command1_Click()
Form11.Show

Unload Form12
End Sub

Private Sub Command2_Click()
rs12.MoveFirst
movefields
End Sub

Private Sub Command3_Click()
rs12.MoveNext
If rs12.EOF Then
rs12.MoveLast
End If
movefields
End Sub

Private Sub Command4_Click()
rs12.MoveLast
movefields
End Sub

Private Sub Command5_Click()
rs12.MovePrevious
If rs12.BOF Then
rs12.MoveFirst
End If
movefields
End Sub

Private Sub delete_Click()
rs12.delete
rs12.MoveNext
If rs.EOF Then
rs12.MoveLast
End If
movefields
End Sub

Private Sub edit_Click()
If rs12.EditMode = dbEditNone Then
rs12.edit
End If
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub movefields()
Text1.Text = rs12("cid")
Text2.Text = rs12("cname")
Text3.Text = rs12("street")
Text4.Text = rs12("city")
Text5.Text = rs12("acchod")
Text6.Text = rs12("pwd")
End Sub

Private Sub save_Click()
rs12("cid") = Text1.Text
rs12("cname") = Text2.Text
rs12("street") = Text3.Text
rs12("city") = Text4.Text
rs12("acchod") = Text5.Text
rs12("pwd") = Text6.Text
rs12.Update
End Sub
