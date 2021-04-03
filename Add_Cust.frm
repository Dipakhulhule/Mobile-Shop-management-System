VERSION 5.00
Begin VB.Form AddCustomer 
   BackColor       =   &H00C000C0&
   Caption         =   "Form2"
   ClientHeight    =   6360
   ClientLeft      =   6285
   ClientTop       =   2640
   ClientWidth     =   4740
   LinkTopic       =   "Form2"
   ScaleHeight     =   6360
   ScaleWidth      =   4740
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Add New Custmer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      Begin VB.OptionButton Option2 
         BackColor       =   &H000000FF&
         Caption         =   "Female"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   3480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Caption         =   "Male"
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
         Left            =   1680
         TabIndex        =   14
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Mobile 
         Height          =   405
         Left            =   1680
         TabIndex        =   10
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox C_name 
         DataField       =   "Custmor_Name"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox A_no 
         DataField       =   "Customer_Aadhar"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox C_Address 
         DataField       =   "Customer_Address"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   4080
         Width           =   1695
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
         Height          =   420
         Left            =   1560
         TabIndex        =   3
         Top             =   5160
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
         Height          =   420
         Left            =   2760
         TabIndex        =   2
         Top             =   5160
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
         Height          =   420
         Left            =   480
         TabIndex        =   1
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Gender 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         TabIndex        =   13
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Aadhar No"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   7
         Top             =   4200
         Width           =   690
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
      Index           =   1
      Left            =   720
      TabIndex        =   11
      Top             =   2880
      Width           =   75
   End
End
Attribute VB_Name = "AddCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command2_Click()
C_name.Text = ""
A_no.Text = ""

mobile.Text = ""
C_Address.Text = ""


End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Save_Click()
Dim Gender As String
    If Option1.Value = True Then
       Gender = "Male"
    ElseIf Option2.Value = True Then
        Gender = "Female"
    End If
     
 A_no = A_no.Text
 
Set con = connect
Set rs = con.Execute("select * from  Customer where Aadhar = '" & A_no & "'")
If rs.EOF Then
con.Execute ("insert into Customer(Name,Aadhar,Gender,Mobile,Address) values('" & C_name.Text & "','" & A_no.Text & "','" & Gender & "','" & mobile.Text & "','" & C_Address.Text & "')")
MsgBox "Record saved"
Else
MsgBox "Custmor Already Exist"
End If
End Sub




