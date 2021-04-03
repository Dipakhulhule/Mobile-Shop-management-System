VERSION 5.00
Begin VB.Form bill 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   6585
   ClientTop       =   3105
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   5040
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      Caption         =   "Bill"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
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
         Left            =   1800
         TabIndex        =   20
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Total 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label price 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   1920
         TabIndex        =   18
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Quantity 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Model 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Brand 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label productname 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label cname 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label date1 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2760
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label bno 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   855
      End
      Begin VB.Label pname 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date And Time"
         Height          =   195
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   2040
         Width           =   735
      End
   End
   Begin VB.Label rem1 
      Caption         =   "Label2"
      Height          =   255
      Left            =   2880
      TabIndex        =   21
      Top             =   5040
      Width           =   975
   End
End
Attribute VB_Name = "bill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Form_Load()
date1.Caption = Now()

cname.Caption = Sell.Combo1.Text
productname.Caption = Sell.product.Text
Model.Caption = Sell.Model.Text
Quantity.Caption = Sell.qunty.Text

Table = productname.Caption





If Sell.product.Text = "Mobiles" Then
Brand.Caption = Sell.Text1.Text
price.Caption = Sell.Text5.Text
p = Sell.Text5.Text


ElseIf Sell.product.Text = "Charger" Then
Brand.Caption = Sell.CT1.Text
price.Caption = Sell.CT7(0).Text
p = Sell.CT7(0).Text

ElseIf Sell.product.Text = "Battery" Then
Brand.Caption = Sell.b7.Text
price.Caption = Sell.b1.Text
p = Sell.b7.Text


ElseIf Sell.product.Text = "Headphone" Then
Brand.Caption = Sell.Brand.Text
price.Caption = Sell.Pi.Text
p = Sell.Pi.Text

End If


Total.Caption = Val(p) * Val(Sell.qunty.Text)

Set con = connect

con.Execute ("insert into sell (Customer_Name,Date_Time,Product_Name,Brand,Model) values('" & cname.Caption & "','" & date1.Caption & "','" & productname.Caption & "','" & Brand.Caption & "','" & Model.Caption & "')")




Set rs = con.Execute("select MAX(Bill_NO)From sell")
 
bno.Caption = rs(0).Value


Set rs = con.Execute(" Select Quantity  From " + Table + " Where Model = '" + Sell.Model.Text + "'  ")

totalqunty = Val(rs(0).Value)
rem1.Caption = totalqunty - Val(Quantity.Caption)

con.Execute ("Update " + Table + " Set Quantity = '" + rem1.Caption + "'  where model ='" + Sell.Model.Text + "' ")

End Sub

