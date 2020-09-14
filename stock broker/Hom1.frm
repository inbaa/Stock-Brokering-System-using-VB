VERSION 5.00
Begin VB.Form Hom1 
   Caption         =   "User Home Page"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu m1 
      Caption         =   "Customer"
   End
   Begin VB.Menu m2 
      Caption         =   "Sales"
   End
End
Attribute VB_Name = "Hom1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub m1_Click()
Me.Hide
Unload Me
Customer.Show

End Sub

Private Sub m2_Click()
Me.Hide
Unload Me
Sales.Show

End Sub
