VERSION 5.00
Begin VB.Form Sell 
   BackColor       =   &H00FF00FF&
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   6750
   ClientTop       =   2340
   ClientWidth     =   6465
   ForeColor       =   &H00C000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   6465
   Visible         =   0   'False
   Begin VB.TextBox qunty 
      Height          =   285
      Left            =   960
      TabIndex        =   75
      Top             =   7560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   960
      TabIndex        =   73
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton sale 
      Caption         =   "Sale"
      Height          =   495
      Left            =   3120
      TabIndex        =   71
      Top             =   6960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Charger 
      BackColor       =   &H000000FF&
      Caption         =   "Charger Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   600
      TabIndex        =   52
      Top             =   3120
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox CT8 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3960
         TabIndex        =   68
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox CT7 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   67
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox CT6 
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   66
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox CT5 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   65
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox CT4 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3960
         TabIndex        =   64
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox CT3 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         TabIndex        =   63
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox CT2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1080
         TabIndex        =   62
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox CT1 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         TabIndex        =   61
         Top             =   720
         Width           =   855
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
         Index           =   7
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   510
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
         Index           =   6
         Left            =   3960
         TabIndex        =   59
         Top             =   1320
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
         Index           =   4
         Left            =   1560
         TabIndex        =   58
         Top             =   1320
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
         Index           =   2
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         Width           =   1095
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
         Index           =   3
         Left            =   1080
         TabIndex        =   56
         Top             =   360
         Width           =   825
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
         Index           =   5
         Left            =   3840
         TabIndex        =   55
         Top             =   360
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
         Index           =   1
         Left            =   2040
         TabIndex        =   54
         Top             =   360
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
         Index           =   2
         Left            =   3120
         TabIndex        =   53
         Top             =   1320
         Width           =   450
      End
   End
   Begin VB.Frame Headphone 
      BackColor       =   &H000000FF&
      Caption         =   "HeadPhone Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   1680
      TabIndex        =   43
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox Brand 
         Height          =   360
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Pi 
         Height          =   405
         Left            =   1320
         TabIndex        =   45
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox m 
         Height          =   375
         Left            =   1200
         TabIndex        =   44
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label15 
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
         Height          =   375
         Left            =   1440
         TabIndex        =   70
         Top             =   240
         Width           =   615
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
         Index           =   2
         Left            =   240
         TabIndex        =   50
         Top             =   1200
         Width           =   435
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
         Index           =   3
         Left            =   240
         TabIndex        =   49
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Index           =   2
         Left            =   2640
         TabIndex        =   48
         Top             =   720
         Width           =   525
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
         Index           =   2
         Left            =   1440
         TabIndex        =   47
         Top             =   1200
         Width           =   495
      End
   End
   Begin VB.Frame Battery 
      BackColor       =   &H000000FF&
      Caption         =   "Battery Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   480
      TabIndex        =   29
      Top             =   3240
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox b7 
         Height          =   285
         Left            =   2880
         TabIndex        =   42
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox b6 
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox b5 
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox b4 
         Height          =   285
         Left            =   3720
         TabIndex        =   39
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox b3 
         Height          =   375
         Left            =   1920
         TabIndex        =   38
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox b2 
         Height          =   375
         Left            =   1080
         TabIndex        =   37
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox b1 
         Height          =   285
         Left            =   240
         TabIndex        =   36
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
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
         Left            =   3000
         TabIndex        =   69
         Top             =   1440
         Width           =   450
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
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   1080
         TabIndex        =   34
         Top             =   360
         Width           =   525
      End
      Begin VB.Label S 
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
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   33
         Top             =   360
         Width           =   1575
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
         Left            =   3720
         TabIndex        =   32
         Top             =   360
         Width           =   1515
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
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   1440
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
         Height          =   435
         Index           =   1
         Left            =   1920
         TabIndex        =   30
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame Mobile 
      BackColor       =   &H000000FF&
      Caption         =   "Mobile_Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.TextBox Text11 
         Height          =   300
         Left            =   4680
         TabIndex        =   28
         Top             =   1920
         Width           =   850
      End
      Begin VB.TextBox Text10 
         Height          =   300
         Left            =   3600
         TabIndex        =   27
         Top             =   1920
         Width           =   850
      End
      Begin VB.TextBox Text9 
         Height          =   300
         Left            =   2280
         TabIndex        =   26
         Top             =   1920
         Width           =   850
      End
      Begin VB.TextBox Text8 
         Height          =   300
         Left            =   1200
         TabIndex        =   25
         Top             =   1920
         Width           =   850
      End
      Begin VB.TextBox Text7 
         Height          =   300
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   850
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Left            =   4800
         TabIndex        =   23
         Top             =   720
         Width           =   850
      End
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   3840
         TabIndex        =   22
         Top             =   720
         Width           =   850
      End
      Begin VB.TextBox Text4 
         Height          =   300
         Left            =   2880
         TabIndex        =   21
         Top             =   720
         Width           =   850
      End
      Begin VB.TextBox Text3 
         Height          =   300
         Left            =   1920
         TabIndex        =   20
         Top             =   720
         Width           =   850
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   960
         TabIndex        =   19
         Top             =   720
         Width           =   850
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   0
         TabIndex        =   18
         Top             =   720
         Width           =   850
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display Type"
         Height          =   195
         Left            =   4560
         TabIndex        =   17
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Screen Size"
         Height          =   255
         Left            =   3480
         TabIndex        =   16
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   255
         Left            =   3960
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Processor"
         Height          =   195
         Left            =   1200
         TabIndex        =   14
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ram"
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Operating System"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rear Camera"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Rom"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Front Camera"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   8
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Brand"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Select  Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton View 
         Caption         =   "View Details"
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
         Left            =   720
         TabIndex        =   6
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox Model 
         Height          =   315
         Left            =   1560
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Product 
         Height          =   315
         ItemData        =   "Sell.frx":0000
         Left            =   1560
         List            =   "Sell.frx":0010
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Model"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select Product"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1275
      End
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Quantity"
      Height          =   255
      Left            =   2760
      TabIndex        =   77
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label label17 
      Height          =   255
      Left            =   4440
      TabIndex        =   76
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Quantity"
      Height          =   255
      Left            =   960
      TabIndex        =   74
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Customer"
      Height          =   255
      Left            =   960
      TabIndex        =   72
      Top             =   6000
      Width           =   1455
   End
End
Attribute VB_Name = "Sell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Combo1_GotFocus()
Set con = connect
sql = "select Name From Customer"

Set rs = con.Execute(sql)
While Not rs.EOF
Combo1.AddItem (rs(0).Value)
rs.MoveNext
 Wend
End Sub

Private Sub Product_GotFocus()
Model.Clear
battery.Visible = False
mobile.Visible = False
Charger.Visible = False
headphone.Visible = False
Combo1.Visible = False
sale.Visible = False

Label16.Visible = False
Label18.Visible = False
qunty.Visible = False
Label19.Visible = False
label17.Visible = False




End Sub

Private Sub Product_LostFocus()
If product.Text = "" Then
     MsgBox "please Select Product ", vbExclamation
     
Else
Set con = connect
Dim Table As String
Table = product.Text


Dim sql As String
sql = "select Model From " + Table

Set rs = con.Execute(sql)
While Not rs.EOF
Model.AddItem (rs(0).Value)
rs.MoveNext
 Wend
End If
End Sub

Private Sub Sell_Click()
bill.Show
End Sub

Private Sub sale_Click()
If Combo1.Text = "" Then
MsgBox "Please Select Customer"
ElseIf label17.Caption = "0" Then
MsgBox "Stock Is Empty"

ElseIf qunty.Text = "" Then
MsgBox "Please Enter Quantity"
ElseIf Val(qunty.Text) > label17.Caption Then
MsgBox "Only " + label17.Caption + " Quantity  Availble"



Else
bill.Show

End If



End Sub

Private Sub View_Click()
If Model.Text = "" Then
MsgBox "please Select Model"
Else
Set con = connect
If product.Text = "Mobiles" Then
mobile.Visible = True
Set rs = con.Execute("SELECT Brand ,model ,Ram,Rom,price, F_cam, R_cam, Processor,os,Screen_size,Display_Type FROM mobiles where model='" & Model.Text & "'")
If Not rs.EOF Then

While Not rs.EOF
Text1.Text = rs(0).Value
Text2.Text = rs(1).Value
Text3.Text = rs(2).Value
Text4.Text = rs(3).Value
Text5.Text = rs(4).Value
Text6.Text = rs(5).Value
Text7.Text = rs(6).Value
Text8.Text = rs(7).Value
Text9.Text = rs(8).Value
Text10.Text = rs(9).Value
Text11.Text = rs(10).Value
rs.MoveNext
Wend
End If


ElseIf product.Text = "Battery" Then
battery.Visible = True
Set rs = con.Execute("SELECT Brand ,Model,Compatible_Brand,Compatible_Model,Capacity, Warranty,Price FROM Battery where model='" & Model.Text & "'")
If Not rs.EOF Then



While Not rs.EOF
b1.Text = rs(0).Value
b2.Text = rs(1).Value
b3.Text = rs(2).Value
b4.Text = rs(3).Value
b5.Text = rs(4).Value
b6.Text = rs(5).Value
b7.Text = rs(6).Value
rs.MoveNext
Wend
End If



ElseIf product.Text = "Headphone" Then
headphone.Visible = True
Set rs = con.Execute("SELECT Brand ,Model,Type,Price FROM Headphone where model='" & Model.Text & "'")
If Not rs.EOF Then

While Not rs.EOF
Brand.Text = rs(0).Value
m.Text = rs(1).Value
Text19.Text = rs(2).Value
Pi.Text = rs(3).Value
rs.MoveNext
Wend
End If




ElseIf product.Text = "Charger" Then
Charger.Visible = True

Set rs = con.Execute("SELECT Brand ,Model,Connector_type,Cable_length,Power_Input, Power_Output,Price,Warranty FROM Charger where model='" & Model.Text & "'")
If Not rs.EOF Then
While Not rs.EOF
CT1.Text = rs(0).Value
CT2.Text = rs(1).Value
CT3.Text = rs(2).Value
CT4.Text = rs(3).Value
CT5.Text = rs(4).Value
CT6(0).Text = rs(5).Value
CT7(1).Text = rs(6).Value
CT8.Text = rs(7).Value
rs.MoveNext
Wend
End If






End If
End If
Dim Table As String
Table = product.Text

Set rs = con.Execute(" Select Quantity  From " + Table + " Where Model = '" + Model.Text + "'  ")
label17.Caption = rs(0).Value














Combo1.Visible = True
sale.Visible = True
Label16.Visible = True
Label18.Visible = True
qunty.Visible = True
Label19.Visible = True
label17.Visible = True


End Sub
