VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   600
      Picture         =   "Form2.frx":15304
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SUBMIT"
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   9960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4560
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   1650
      Left            =   -120
      Picture         =   "Form2.frx":18F1C
      ScaleHeight     =   1590
      ScaleWidth      =   20595
      TabIndex        =   0
      Top             =   -480
      Width           =   20655
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Text1.SetFocus
    If Text1.Text = "anand" Then
        Form11.Show
        Else
        MsgBox "PASSWORD IS WRONG"
            End If
            Unload Form2
            
End Sub

Private Sub Command2_Click()
Form1.Show

Unload Form2
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Picture2_Click()
Form1.Show

Unload Form2
End Sub
