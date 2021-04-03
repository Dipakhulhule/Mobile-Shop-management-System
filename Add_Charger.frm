VERSION 5.00
Begin VB.Form AddC 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   6585
   ClientTop       =   2490
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   4800
   Begin VB.Frame a 
      BackColor       =   &H000000FF&
      Caption         =   "Add Charger Details"
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.TextBox Output 
         Height          =   405
         Left            =   1800
         TabIndex        =   11
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox ModelNo 
         Height          =   400
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox Brand 
         Height          =   400
         Left            =   1800
         TabIndex        =   9
         Top             =   600
         Width           =   1605
      End
      Begin VB.ComboBox C_length 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox Conn_Type 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox price 
         Height          =   400
         Left            =   1800
         TabIndex        =   6
         Top             =   4320
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
         Height          =   480
         Left            =   120
         TabIndex        =   5
         Top             =   6120
         Width           =   1215
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
         Height          =   480
         Left            =   1440
         TabIndex        =   4
         Top             =   6120
         Width           =   1095
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
         Height          =   480
         Left            =   2760
         TabIndex        =   3
         Top             =   6120
         Width           =   975
      End
      Begin VB.ComboBox Warranty 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   4920
         Width           =   1455
      End
      Begin VB.TextBox pinput 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label S 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
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
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label S 
         BackStyle       =   0  'Transparent
         Caption         =   "Warranty"
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
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " power Ouput"
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
         Left            =   120
         TabIndex        =   17
         Top             =   3840
         Width           =   1155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Power Input"
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
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Model No"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label S 
         BackStyle       =   0  'Transparent
         Caption         =   "Cable length"
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
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label S 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connector Type"
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
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   4440
         Width           =   450
      End
   End
End
Attribute VB_Name = "AddC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Brand.Text = ""
ModelNo.Text = ""
Conn_Type.Text = ""
C_length.Text = ""
pinput.Text = ""
Output.Text = ""
price.Text = ""
Warranty.Text = ""

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Save_Click()
Q = "0"
t1 = ModelNo.Text
Set con = connect
Set con = connect
con.Execute ("insert into Charger(Brand,model,Connector_type,Cable_length,Power_input,Power_output,Price,Warranty,Quantity) values('" & Brand.Text & "','" & ModelNo.Text & "','" & Conn_Type.Text & "','" & C_length.Text & "','" & pinput.Text & "','" & Output.Text & "','" & price.Text & "','" & Warranty.Text & "','" & Q & "')")
MsgBox "Record saved"
con.Close


End Sub

