VERSION 5.00
Begin VB.Form Add_Accessories 
   ClientHeight    =   7755
   ClientLeft      =   6690
   ClientTop       =   2955
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   5415
   Begin VB.Frame Frame1 
      Caption         =   "Add Accessories"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox Quantity 
         Height          =   450
         Left            =   2520
         TabIndex        =   15
         Top             =   4800
         Width           =   1005
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   6360
         Width           =   735
      End
      Begin VB.CommandButton Reset 
         Caption         =   "Reset"
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   6360
         Width           =   735
      End
      Begin VB.CommandButton Save 
         Caption         =   "Save"
         Height          =   495
         Left            =   600
         TabIndex        =   12
         Top             =   6360
         Width           =   855
      End
      Begin VB.ComboBox Accessorises 
         Height          =   315
         Left            =   2520
         TabIndex        =   11
         Top             =   3480
         Width           =   1000
      End
      Begin VB.ComboBox Supplier 
         Height          =   315
         Left            =   2520
         TabIndex        =   10
         Top             =   5640
         Width           =   1000
      End
      Begin VB.TextBox price 
         Height          =   450
         Left            =   2520
         TabIndex        =   9
         Top             =   4080
         Width           =   1000
      End
      Begin VB.TextBox C_Name 
         Height          =   450
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Model 
         Height          =   450
         Left            =   2520
         TabIndex        =   7
         Top             =   2640
         Width           =   1000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   4920
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Select Supplier "
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   5640
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Price"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   5
         Top             =   4320
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Accesoories"
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "Model"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Company Name"
         Height          =   195
         Left            =   600
         TabIndex        =   2
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enter id"
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   960
         Width           =   540
      End
   End
End
Attribute VB_Name = "Add_Accessories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Save_Click()
Dim con As New ADODB.Connection
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= D:\Mobile_Shop_mgt_System\Database\DMS.mdb"
con.Open
con.Execute ("insert into Accessories values('" & B_id.Text & "','" & Model.Text & "','" & C_Name.Text & "','" & Warranty.Text & "','" & price.Text & "')")
MsgBox "Record saved"
con.Close
End Sub


