VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   855
   ClientWidth     =   4560
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3.18557e8
   ScaleMode       =   0  'User
   ScaleWidth      =   3.0476e9
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   -240
      Picture         =   "Form1.frx":B5685
      ScaleHeight     =   1755
      ScaleMode       =   0  'User
      ScaleWidth      =   20600
      TabIndex        =   3
      Top             =   -600
      Width           =   20655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Caption         =   "SELECT USER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   6360
      TabIndex        =   0
      Top             =   4920
      Width           =   7215
      Begin VB.CommandButton Command2 
         BackColor       =   &H00004040&
         Caption         =   "Account Holder"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00004040&
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   960
         TabIndex        =   1
         Top             =   1320
         Width           =   2295
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show

Unload Form1
End Sub

Private Sub Command2_Click()
Form3.Show

Unload Form1
End Sub

Private Sub exit_Click()
End
End Sub

