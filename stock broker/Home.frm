VERSION 5.00
Begin VB.Form Hom 
   Caption         =   "Admin Home Page"
   ClientHeight    =   6825
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data4 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Data Data3 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sales"
      Top             =   4920
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "production_plan"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   0
      Picture         =   "Home.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
   Begin VB.Menu m1 
      Caption         =   "Supplier "
   End
   Begin VB.Menu m2 
      Caption         =   "Purchase stock"
   End
   Begin VB.Menu m3 
      Caption         =   "Supplier stock"
   End
   Begin VB.Menu m4 
      Caption         =   "Purchase"
   End
   Begin VB.Menu m10 
      Caption         =   "Stock arrival"
   End
   Begin VB.Menu m5 
      Caption         =   "Sales stock"
   End
   Begin VB.Menu m6 
      Caption         =   "Production plan"
   End
   Begin VB.Menu m7 
      Caption         =   "Production completion"
   End
   Begin VB.Menu m8 
      Caption         =   "Customer"
   End
   Begin VB.Menu m9 
      Caption         =   "Sales"
   End
   Begin VB.Menu m16 
      Caption         =   "Report"
      Begin VB.Menu m11 
         Caption         =   "Pending production plan"
      End
      Begin VB.Menu m12 
         Caption         =   "Pending purchase "
      End
      Begin VB.Menu m13 
         Caption         =   "Purchase"
      End
      Begin VB.Menu m14 
         Caption         =   "Sales"
      End
      Begin VB.Menu m15 
         Caption         =   "Stock"
      End
   End
End
Attribute VB_Name = "Hom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Customer_Click()

End Sub

Private Sub Production_complete_Click()

End Sub

Private Sub Production_plan_Click()

End Sub

Private Sub Purchase_Click()

End Sub


Private Sub m1_Click()
Me.Hide
Unload Me
Supplier.Show
End Sub

Private Sub m10_Click()
On Error GoTo e
Me.Hide
Stock_arrival.Show
'vbModal, Me
'Unload Me
Exit Sub
e:
Resume Next
End Sub

Private Sub m11_Click()
Data2.Refresh
Report_Pending_production_plan.Show
End Sub

Private Sub m12_Click()
Data1.Refresh
Report_Purchase_pending.Show
End Sub

Private Sub m13_Click()
Data1.Refresh
Report_purchase.Show
End Sub

Private Sub m14_Click()
Data3.Refresh
Report_sales.Show
End Sub

Private Sub m15_Click()
Data4.Refresh
Report_stock.Show
End Sub

Private Sub m2_Click()
Me.Hide
Unload Me
Stock.Show
End Sub

Private Sub m3_Click()
Me.Hide
Unload Me
Supplier_Stock.Show
End Sub

Private Sub m4_Click()
On Error GoTo e
Me.Hide
Purchase.Show
'vbModal, Me
'Unload Me
Exit Sub
e:
Resume Next
End Sub

Private Sub m5_Click()
Me.Hide
Unload Me
Sale_stock.Show
End Sub

Private Sub m6_Click()
Me.Hide
Unload Me
Production_plan.Show
End Sub

Private Sub m7_Click()
Me.Hide
Unload Me
Production_complete.Show
End Sub

Private Sub m8_Click()
Me.Hide
Unload Me
Customer.Show
End Sub

Private Sub m9_Click()
Me.Hide
Unload Me
Sales.Show
End Sub
