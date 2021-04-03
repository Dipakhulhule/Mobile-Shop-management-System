VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00FF00FF&
   Caption         =   "Mobile Shop"
   ClientHeight    =   3465
   ClientLeft      =   7050
   ClientTop       =   4050
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   4350
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Login"
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "Forgot Password"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   2520
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Login 
         Caption         =   "Login"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Pass 
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox uname 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End

End Sub

Private Sub Login_Click()
Set con = connect


Set rs = con.Execute("Select * From Login Where Username = '" + uname.Text + "' and Password = '" + Pass.Text + "'")

If rs.EOF Then
MsgBox "Wrong Username and password"
Else
Main.Show

End If

End Sub
