VERSION 5.00
Begin VB.Form Add_mobile 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   5970
   ClientTop       =   2175
   ClientWidth     =   5565
   FillColor       =   &H00FF80FF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   5565
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Add New Mobile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.TextBox Brand 
         Height          =   405
         Left            =   1200
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Exit 
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
         Left            =   3240
         TabIndex        =   25
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton Reset 
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
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   6120
         Width           =   855
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
         Height          =   375
         Left            =   1200
         TabIndex        =   23
         Top             =   6120
         Width           =   735
      End
      Begin VB.TextBox Display_type 
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox s_size 
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox os 
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Processor 
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox R_cam 
         Height          =   375
         Left            =   3840
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox F_Cam 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox Price 
         Height          =   405
         Left            =   1200
         TabIndex        =   11
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox Rom 
         Height          =   495
         Left            =   1200
         TabIndex        =   10
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox Ram 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Model 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display type"
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
         Left            =   2400
         TabIndex        =   22
         Top             =   3840
         Width           =   1050
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Size"
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
         Left            =   2400
         TabIndex        =   21
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "operating System"
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
         Left            =   2280
         TabIndex        =   20
         Top             =   2280
         Width           =   1470
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proceesorr"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rear camera"
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
         Left            =   2520
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Front Camera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   4800
         Width           =   1095
      End
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ram"
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
         TabIndex        =   3
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label2 
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
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
End
Attribute VB_Name = "Add_mobile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Exit_Click()
Unload Me

End Sub

Private Sub Reset_Click()
Brand.Text = ""
Model.Text = ""
Ram.Text = ""
Rom.Text = ""
price.Text = ""
R_cam.Text = ""
F_Cam.Text = ""
Processor.Text = ""
os.Text = ""
s_size.Text = ""
Display_type.Text = ""



End Sub

Private Sub Save_Click()
Q = "0"
Set con = connect
Set rs = con.Execute("select * from  Mobiles where Model = '" & Model.Text & "'")
If rs.EOF Then
con.Execute ("insert into Mobiles(Brand,model,Ram,Rom,price,F_cam,R_Cam,Processor,os,Screen_size,Display_type,Quantity) values('" & Brand.Text & "','" & Model.Text & "','" & Ram.Text & "','" & Rom.Text & "','" & price.Text & "','" & F_Cam.Text & "','" & R_cam.Text & "','" & Processor.Text & "','" & os.Text & "','" & s_size.Text & "','" & Display_type.Text & "','" & Q & "')")
MsgBox "Record saved"
con.Close
Else
MsgBox "Already Exist"
End If

End Sub

