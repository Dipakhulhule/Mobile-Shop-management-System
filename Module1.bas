Attribute VB_Name = "Module1"

Public Function connect() As Connection
 Dim con As ADODB.Connection

 Set con = New ADODB.Connection

 con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= E:\SEM 5\PROJECT\Mobile_Shop_mgt_System\Database\DMS.mdb"
 Set connect = con
 End Function
 
 
 
 
 
