VERSION 5.00
Begin VB.Form Supplier1 
   BackColor       =   &H00FF00FF&
   Caption         =   "Add Supplier"
   ClientHeight    =   4845
   ClientLeft      =   7365
   ClientTop       =   2175
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4065
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Add New Supplier"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.TextBox name1 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox A_no 
         Height          =   405
         Left            =   1440
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox mobile 
         Height          =   360
         Left            =   1440
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Address 
         Height          =   360
         Left            =   1440
         TabIndex        =   4
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Save 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   2
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   1
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Aadhar No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   555
      End
   End
End
Attribute VB_Name = "Supplier1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
name1.Text = ""
mobile.Text = ""
A_no.Text = ""
Address.Text = ""

End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Save_Click()
 
 Set con = connect
 Set rs = con.Execute("select * from  Supplier where Aadhar = '" & A_no.Text & "'")
If rs.EOF Then
con.Execute ("insert into Supplier(S_Name,Mobile,Aadhar,Address) values('" & name1.Text & "','" & mobile.Text & "','" & A_no.Text & "','" & Address.Text & "')")
MsgBox "Record saved"
Else
MsgBox "Supplier already Exist"
End If
End Sub
