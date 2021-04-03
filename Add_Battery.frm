VERSION 5.00
Begin VB.Form Add_Battery 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF80FF&
   Caption         =   "Add Battery Details"
   ClientHeight    =   6825
   ClientLeft      =   7050
   ClientTop       =   2340
   ClientWidth     =   4350
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H000080FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   4350
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Add Battery Details"
      Height          =   6255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3855
      Begin VB.ComboBox Warranty 
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   3840
         Width           =   1455
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
         TabIndex        =   10
         Top             =   5160
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
         Height          =   480
         Left            =   1560
         TabIndex        =   9
         Top             =   5160
         Width           =   1095
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
         TabIndex        =   8
         Top             =   5160
         Width           =   1215
      End
      Begin VB.TextBox price 
         Height          =   400
         Left            =   1800
         TabIndex        =   7
         Top             =   4320
         Width           =   1455
      End
      Begin VB.ComboBox Compatible_brand 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox Compatible_Model 
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Brand 
         Height          =   400
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox ModelNo 
         Height          =   400
         Left            =   1800
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Capacity 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   17
         Top             =   4320
         Width           =   450
      End
      Begin VB.Label S 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compatible Brand"
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
         TabIndex        =   16
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label S 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   510
      End
      Begin VB.Label S 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compatible Model"
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
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacity"
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
         Left            =   480
         TabIndex        =   12
         Top             =   3360
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "warranty"
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
         Left            =   480
         TabIndex        =   11
         Top             =   3840
         Width           =   735
      End
   End
   Begin VB.Label S 
      Caption         =   "Supplier id"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "Add_Battery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Brand.Text = ""
ModelNo.Text = ""
Compatible_brand.Text = ""
Capacity.Text = ""
Warranty.Text = ""
price.Text = ""


End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Save_Click()
Q = "0"
t1 = ModelNo.Text
Set con = connect
Set rs = con.Execute("select * from  Battery where Model = '" & t1 & "'")
If rs.EOF Then
con.Execute ("insert into Battery(Brand,Model,Compatible_Brand,Compatible_Model,Capacity,Warranty,price,Quantity) values('" & Brand.Text & "','" & ModelNo.Text & "','" & Compatible_brand.Text & "','" & Compatible_Model.Text & "','" & Capacity.Text & "','" & Warranty.Text & "','" & price.Text & "','" & Q & "')")
MsgBox "Record saved"
con.Close
Else
MsgBox "Model Already  Exist"
End If
End Sub




