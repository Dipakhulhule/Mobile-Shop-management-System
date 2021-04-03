VERSION 5.00
Begin VB.Form AddHeadphone 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   6285
   ClientTop       =   2955
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   4740
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Add New Headphone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      Begin VB.TextBox m 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Pi 
         Height          =   405
         Left            =   1680
         TabIndex        =   11
         Top             =   2880
         Width           =   1575
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
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   4200
         Width           =   855
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
         Left            =   1680
         TabIndex        =   5
         Top             =   4200
         Width           =   735
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
         Left            =   480
         TabIndex        =   4
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Brand 
         Height          =   360
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox Type1 
         Height          =   315
         ItemData        =   "AddHeadphone.frx":0000
         Left            =   1680
         List            =   "AddHeadphone.frx":000A
         TabIndex        =   2
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Brand"
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
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Width           =   570
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   435
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Left            =   1080
      TabIndex        =   0
      Top             =   3720
      Width           =   75
   End
End
Attribute VB_Name = "AddHeadphone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Save_Click()

Q = "0"
Set con = connect

Set rs = con.Execute("select * from  Headphone where Model = '" & m.Text & "'")
If rs.EOF Then
con.Execute ("insert into Headphone(Brand,Model,Type,Price,Quantity) values('" & Brand.Text & "','" & m.Text & "','" & Type1.Text & "','" & Pi.Text & "','" & Q & "')")
MsgBox "Record saved"
Else
MsgBox "Model Already Exist"
End If

con.Close
End Sub
