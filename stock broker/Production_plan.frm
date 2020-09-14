VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Production_plan 
   Caption         =   "Production Plan"
   ClientHeight    =   8085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "GO TO PRODUCTION COMPLETE FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   29
      Top             =   6720
      Width           =   3975
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sale_stock"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      DataField       =   "house_no"
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H0080FFFF&
      TabIndex        =   10
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H0080FFFF&
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7200
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Production_plan.frx":0000
      Left            =   3240
      List            =   "Production_plan.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "house_no"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "supp_name"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3240
      TabIndex        =   9
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   240
      TabIndex        =   19
      Top             =   360
      Width           =   6375
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   20
         Top             =   480
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   136511489
         CurrentDate     =   41058
      End
      Begin VB.Label Label2 
         Caption         =   "Purchase Item Number"
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
         Left            =   360
         TabIndex        =   28
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label20 
         Caption         =   "Date"
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
         Left            =   360
         TabIndex        =   26
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Sale Item Name"
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
         Left            =   360
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Sale Item Number"
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
         Left            =   360
         TabIndex        =   24
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Production Plan Number"
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
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "@"
         Height          =   255
         Left            =   4800
         TabIndex        =   22
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Quantity to be Produced"
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
         Left            =   360
         TabIndex        =   21
         Top             =   3000
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000017&
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   5400
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "ADD NEW PLAN"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   6960
      TabIndex        =   15
      Top             =   360
      Width           =   2295
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Navigate to record"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   6960
      TabIndex        =   14
      Top             =   4200
      Width           =   2295
      Begin VB.CommandButton Command9 
         Caption         =   "CLEAR"
         Height          =   615
         Left            =   480
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "production_plan"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label21 
      Caption         =   "Production Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3120
      TabIndex        =   27
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Production_plan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num As Integer
Public db1 As Database
Public rs1 As Recordset
Public db As Database
Public rs As Recordset

Private Sub Command1_Click()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 DTPicker1.Value = Date
Combo1.ListIndex = -1
'AUTO GENERATE
 Data1.Refresh
 v = Data1.Recordset.RecordCount
 If v = 0 Then
  Text1.Text = "Pplan_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 7)
  res = num + 1
  Text1.Text = "Pplan_" & res
  End If
  Command12.Enabled = False
  'edit,delete and update
 Command2.Enabled = False
 'Command3.Enabled = False
 Command4.Enabled = False
 'navigation frame
 Frame3.Enabled = False
 Command5.Enabled = False
 Command6.Enabled = False
 Command7.Enabled = False
 Command8.Enabled = False
 'SAVE
 Command10.Enabled = True
Command1.Enabled = False
 Command11.Enabled = False
End Sub

Private Sub Command10_Click()
f = 0
z = 0
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
f = 1
ElseIf Not (DTPicker1.Value = Date) Then
MsgBox ("Enter today's date")
f = 1
ElseIf Val(Text4.Text) <= 0 Then
MsgBox ("Quantity cannot be less than 0")
Text4.Text = ""
f = 1
'to cal quantity to be deducted from purchase_stock
ElseIf z = 0 Then
 Set db1 = OpenDatabase("e:\\III year Project\PROJECT.MDB")
 Set rs1 = db1.OpenRecordset("select sur_name1, sur_name2 from sale_stock where sitem_no = '" & Combo1.Text & "' ")
 rs1.MoveFirst
 '100
 n = rs1.Fields(0).Value
 'ml
 uom = rs1.Fields(1).Value
 'quantity to be produced
 q = Val(Text4.Text)
 Set rs1 = db1.OpenRecordset("select quantity,unit_of_measurement,stock_in_progress from purchase_stock where pitem_no = '" & Text3.Text & "' ")
 rs1.MoveFirst
 base_quantity = rs1.Fields(0).Value
 base_uom = rs1.Fields(1).Value
 sip = rs1.Fields(2).Value
 If uom = "ml" Then
 totalq = (n * q) / 1000
 ElseIf uom = "mg" Then
 totalq = (n * q) / 1000
 ElseIf uom = "litres" Then
 totalq = n * q
 ElseIf uom = "kilogram" Then
 totalq = n * q
 End If
 'if unit of measurement is ton in purchase stock
 If base_uom = "ton" Then
 totalq = totalq / 1000
 End If
 'to check if stock is available
 If (base_quantity - totalq) > 0 Then
 base_quantity = base_quantity - totalq
  sip = sip + totalq
 Else
 MsgBox ("Raw material is not available for given quantity")
 Text4.Text = ""
 f = 1
 End If
 
End If

If f = 0 Then
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = DTPicker1.Value
Data1.Recordset.Fields(2) = Combo1.Text
Data1.Recordset.Fields(3) = Text2.Text
Data1.Recordset.Fields(4) = Text3.Text
Data1.Recordset.Fields(5) = Val(Text4.Text)
Data1.Recordset.Fields(6) = "Pending"
Data1.Recordset.Fields(7) = Val(Text4.Text)
'to reduce quantity in purchase stock
Data2.Refresh
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
 If (Text3.Text = Data2.Recordset.Fields(0)) Then
Data2.Recordset.Edit
Data2.Recordset.Fields(2) = base_quantity
Data2.Recordset.Fields(5) = sip
Data2.Recordset.Update
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop
Data1.Recordset.Update
MsgBox ("Plan added successfully")
'Navigation frame
Frame3.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
'edit,delete
Command2.Enabled = True
'Command3.Enabled = True
Command12.Enabled = True
'SELF
Command10.Enabled = False
Command1.Enabled = True
 Command11.Enabled = True
End If

End Sub

Private Sub Command11_Click()
Me.Hide
Unload Me
Hom.Show
End Sub

Private Sub Command12_Click()
On Error GoTo e
Production_complete.Show
Me.Hide
Exit Sub
e:
Resume Next
End Sub

Private Sub Command2_Click()
a = InputBox("Enter Production Plan Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
  If Data1.Recordset.Fields(6) = "Completed" Then
  flag = 3
  Exit Do
  ElseIf Not Data1.Recordset.Fields(5) = Data1.Recordset.Fields(7) Then
  flag = 4
  Exit Do
  Else
  d = Data1.Recordset.Fields(1)
  da = Date
   If StrComp(d, da) = 0 Then
  Text1.Text = Data1.Recordset.Fields(0)
  DTPicker1.Value = Data1.Recordset.Fields(1)
  Combo1.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
  DTPicker1.Enabled = False
  Combo1.Enabled = False
  Text2.Enabled = False
  Text3.Enabled = False
  Text4.Text = Data1.Recordset.Fields(5)
  num = Data1.Recordset.Fields(5)
   flag = 1
  Exit Do
  Else
  flag = 2
  Exit Do
   End If
 End If
 Else
 Data1.Recordset.MoveNext
End If
Loop
If flag = 1 Then
Data1.Recordset.Edit
Command4.Enabled = True
Frame3.Enabled = False
Frame4.Enabled = False
Command1.Enabled = False
'Command3.Enabled = False
Command12.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
ElseIf flag = 2 Then
MsgBox ("Enter today's Production plan number")
ElseIf flag = 3 Then
MsgBox ("This Production plan is completed")
ElseIf flag = 4 Then
MsgBox ("Production has started cannot edit")

Else
MsgBox ("Record not found")
End If

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
z = 0
If Text4.Text = "" Then
a = MsgBox("Please enter quantity!", vbExclamation)
z = 1
ElseIf Val(Text4.Text) <= 0 Then
MsgBox ("Quantity cannot be less than 0")
Text4.Text = ""
z = 1
Else
 Set db1 = OpenDatabase("e:\\III year Project\PROJECT.MDB")
 Set rs1 = db1.OpenRecordset("select sur_name1, sur_name2 from sale_stock where sitem_no = '" & Combo1.Text & "' ")
 rs1.MoveFirst
 '100
 n = rs1.Fields(0).Value
 'ml
 uom = rs1.Fields(1).Value
 'new quantity
 q = Val(Text4.Text)
 Set rs1 = db1.OpenRecordset("select quantity,unit_of_measurement,stock_in_progress from purchase_stock where pitem_no = '" & Text3.Text & "' ")
 rs1.MoveFirst
 base_quantity = rs1.Fields(0).Value
 base_uom = rs1.Fields(1).Value
 sip = rs1.Fields(2).Value
 ' to calc difference in quantity newly entered
 If q < num Then
 'new q < old q (+)
 num = num - q
 f = 1
 End If
 If q > num Then
 'new q> old q (-)
 num = q - num
 f = 2
 End If
 If q = num Then
 'no change in quantity then no update
 z = 1
 End If
 '
 If uom = "ml" Then
 totalq = (n * num) / 1000
 ElseIf uom = "mg" Then
 totalq = (n * num) / 1000
 ElseIf uom = "litres" Then
 totalq = n * num
 ElseIf uom = "kilogram" Then
 totalq = n * num
 End If
 'if unit of measurement is ton in purchase stock
 If base_uom = "ton" Then
 totalq = totalq / 1000
 End If
 'to check if stock is available
If f = 2 Then
 If (base_quantity - totalq) > 0 Then
 base_quantity = base_quantity - totalq
  sip = sip + totalq
 Else
 MsgBox ("Raw material is not available for given quantity")
 Text4.Text = ""
 End If
End If
If f = 1 Then
 base_quantity = base_quantity + totalq
  sip = sip - totalq
End If
 
End If
 If z = 0 Then
 MsgBox (2)
Data1.Recordset.Fields(5) = Val(Text4.Text)
Data1.Recordset.Fields(7) = Val(Text4.Text)
'to reduce quantity in purchase stock
Data2.Refresh
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
 If (Text3.Text = Data2.Recordset.Fields(0)) Then
Data2.Recordset.Edit
Data2.Recordset.Fields(2) = base_quantity
Data2.Recordset.Fields(5) = sip
Data2.Recordset.Update
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop

Data1.Recordset.Update
MsgBox ("Updated successfully")
  'disabled in edit command
  DTPicker1.Enabled = False
  Combo1.Enabled = False
  Text2.Enabled = False
  Text3.Enabled = False
   
Frame3.Enabled = True
Frame4.Enabled = True
Command1.Enabled = True
Command12.Enabled = True
'Command3.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
'SELF
Command4.Enabled = False
End If
End Sub

Private Sub Command5_Click()
Data1.Recordset.MoveFirst
 If Data1.Recordset.BOF = True Then
MsgBox ("No record to display")
Command5.Enabled = False
Else
 Text1.Text = Data1.Recordset.Fields(0)
  DTPicker1.Value = Data1.Recordset.Fields(1)
  Combo1.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
  Text4.Text = Data1.Recordset.Fields(5)
  Command6.Enabled = False
 Command7.Enabled = True
 End If

End Sub

Private Sub Command6_Click()
Data1.Recordset.MovePrevious
If Data1.Recordset.BOF = True Then
MsgBox ("This is first record")
Command6.Enabled = False
Else
 Text1.Text = Data1.Recordset.Fields(0)
  DTPicker1.Value = Data1.Recordset.Fields(1)
  Combo1.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
  Text4.Text = Data1.Recordset.Fields(5)
 Command7.Enabled = True
 End If

End Sub

Private Sub Command7_Click()
Data1.Recordset.MoveNext
If (Data1.Recordset.EOF = True) Then
 MsgBox ("This is last record")
 Command7.Enabled = False
 Else
 Text1.Text = Data1.Recordset.Fields(0)
  DTPicker1.Value = Data1.Recordset.Fields(1)
  Combo1.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
  Text4.Text = Data1.Recordset.Fields(5)
 Command6.Enabled = True
 End If

End Sub

Private Sub Command8_Click()
 Data1.Recordset.MoveLast
 If Data1.Recordset.BOF = True Then
MsgBox ("No record to display")
Command8.Enabled = False
Else
 Text1.Text = Data1.Recordset.Fields(0)
  DTPicker1.Value = Data1.Recordset.Fields(1)
  Combo1.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
  Text4.Text = Data1.Recordset.Fields(5)
 Command7.Enabled = False
 Command6.Enabled = True
 End If

End Sub

Private Sub Command9_Click()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 DTPicker1.Value = Date
Combo1.ListIndex = -1
'NAvigation frame
v = Data1.Recordset.RecordCount
If v = 0 Then
Frame3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Else
Frame3.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
End If
'edit,delete
Command2.Enabled = True
'Command3.Enabled = True
Command12.Enabled = True
'save
Command10.Enabled = False
Command1.Enabled = True
 Command11.Enabled = True
End Sub

Private Sub Form_Activate()
DTPicker1.Value = Date
'add items to combo
Combo1.Clear
If Data3.Recordset.RecordCount <> 0 Then
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
Combo1.AddItem (Data3.Recordset.Fields(0))
Data3.Recordset.MoveNext
Loop
Else
MsgBox ("Sale stock table is empty")
Hom.Show
Unload Me
End If
'edit,delete
Data1.Refresh
v = Data1.Recordset.RecordCount
If v = 0 Then
Command2.Enabled = False
'Command3.Enabled = False
'navigation frame
Frame3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End If
Command10.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Text4_GotFocus()
'adding value to text
'sql
If Combo1.Text = "" Then
MsgBox ("Enter Sale Item Number first")
Else
Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
Set rs = db.OpenRecordset("select sitem_name, pitem_no from sale_stock where sitem_no = '" & Combo1.Text & "' ")
rs.MoveFirst
Text2.Text = rs.Fields(0).Value
Text3.Text = rs.Fields(1).Value
End If

End Sub
