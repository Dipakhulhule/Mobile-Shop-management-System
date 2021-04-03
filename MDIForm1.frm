VERSION 5.00
Begin VB.MDIForm Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000002&
   Caption         =   "Mobile Shop"
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12435
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Entries 
      Caption         =   "Entries"
      Begin VB.Menu Cust 
         Caption         =   " Customer"
      End
      Begin VB.Menu Supp 
         Caption         =   "Supplier"
      End
      Begin VB.Menu product 
         Caption         =   "Product Entries"
         Begin VB.Menu AddMob 
            Caption         =   "Add New Mobile"
         End
         Begin VB.Menu AddHeadphhone 
            Caption         =   "Add New HeadPhone"
         End
         Begin VB.Menu AddCharger 
            Caption         =   "Add new Charger"
         End
         Begin VB.Menu AddBattery 
            Caption         =   "Add New Battery"
         End
      End
   End
   Begin VB.Menu SaleProduct 
      Caption         =   "Sell  Product"
   End
   Begin VB.Menu pur 
      Caption         =   "Purchase Product"
   End
   Begin VB.Menu Report 
      Caption         =   "Report"
      Begin VB.Menu Supplier 
         Caption         =   "Supplier Report"
      End
      Begin VB.Menu Customer 
         Caption         =   "Customer Report"
      End
      Begin VB.Menu p 
         Caption         =   "Purchase Report"
      End
      Begin VB.Menu sells 
         Caption         =   "Sells Report"
      End
      Begin VB.Menu Stock 
         Caption         =   "Stock Report"
         Begin VB.Menu mobile 
            Caption         =   "Mobile"
         End
         Begin VB.Menu Charger 
            Caption         =   "Charger"
            Checked         =   -1  'True
         End
         Begin VB.Menu battery 
            Caption         =   "Battetry"
         End
         Begin VB.Menu headphone 
            Caption         =   "Headphone"
         End
      End
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub AddBattery_Click()
Add_Battery.Show
End Sub

Private Sub AddCharger_Click()
AddC.Show
End Sub

Private Sub AddHeadphhone_Click()
AddHeadphone.Show
End Sub

Private Sub AddMob_Click()
Add_mobile.Show
End Sub

Private Sub AddStock_Click()
Add_Stock.Show
End Sub

Private Sub battery_Click()
BatteryStock.Show

End Sub

Private Sub Charger_Click()
ChargerStock.Show
End Sub

Private Sub cust_Click()
AddCustomer.Show
End Sub

Private Sub Customer_Click()
CustomerReport.Show

End Sub

Private Sub Exit_Click()



End

End Sub

Private Sub headphone_Click()
HeadphoneStock.Show
End Sub

Private Sub mobile_Click()
MobStock.Show
End Sub


Private Sub p_Click()
PurchaseReport.Show
End Sub

Private Sub pur_Click()
Purchase.Show
End Sub

Private Sub SaleProduct_Click()
Sell.Show
End Sub

Private Sub sells_Click()
SellsReport.Show
End Sub

Private Sub Supp_Click()
Supplier1.Show
End Sub

Private Sub Supplier_Click()
SupplierReport.Show
End Sub
