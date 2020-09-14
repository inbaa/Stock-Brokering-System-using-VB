VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Production_complete 
   Caption         =   " Production Complete"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   9690
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      DataField       =   "house_no"
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   4440
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "house_no"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   6
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "house_no"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   5
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "house_no"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "supp_name"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Production_complete.frx":0000
      Left            =   3480
      List            =   "Production_complete.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7560
      MaskColor       =   &H0080FFFF&
      TabIndex        =   10
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   7560
      MaskColor       =   &H0080FFFF&
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7320
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7320
      TabIndex        =   12
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "production_plan"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "production_complete"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   3015
      Left            =   7080
      TabIndex        =   27
      Top             =   4320
      Width           =   2295
      Begin VB.CommandButton Command9 
         Caption         =   "CLEAR"
         Height          =   615
         Left            =   480
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   7080
      TabIndex        =   24
      Top             =   480
      Width           =   2295
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   25
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
         TabIndex        =   26
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000017&
      Height          =   1095
      Left            =   360
      TabIndex        =   23
      Top             =   5760
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "NEW PLAN COMPLETION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   0
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   240
      TabIndex        =   14
      Top             =   480
      Width           =   6615
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3240
         TabIndex        =   2
         Top             =   1320
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   136708097
         CurrentDate     =   41058
      End
      Begin VB.Label Label6 
         Caption         =   "Production Completion Number"
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
         TabIndex        =   30
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity Produced"
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
         TabIndex        =   29
         Top             =   3960
         Width           =   2055
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
         TabIndex        =   22
         Top             =   3000
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
         TabIndex        =   21
         Top             =   1440
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
         TabIndex        =   20
         Top             =   2520
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
         TabIndex        =   19
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Production Plan Number"
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
         TabIndex        =   18
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "@"
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Quantity Pending"
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
         TabIndex        =   16
         Top             =   3480
         Width           =   2175
      End
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sale_stock"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label21 
      Caption         =   "Production Completion Entry"
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
      Left            =   2040
      TabIndex        =   28
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Production_complete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As Database
Public rs As Recordset

Private Sub Command1_Click()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Text5.Text = ""
 Text6.Text = ""
 DTPicker1.Value = Date
Combo1.ListIndex = -1
'add items to combo
Combo1.Clear
If Data2.Recordset.RecordCount <> 0 Then
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If Data2.Recordset.Fields(6) = "Pending" Then
Combo1.AddItem (Data2.Recordset.Fields(0))
Data2.Recordset.MoveNext
Else
Data2.Recordset.MoveNext
End If
Loop
Else
MsgBox ("No pendings in Production plan")
Unload Me
Load Hom
Hom.Show
End If

'AUTO GENERATE
 Data1.Refresh
 v = Data1.Recordset.RecordCount
 If v = 0 Then
  Text1.Text = "Pcomplete_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 11)
  res = num + 1
  Text1.Text = "Pcomplete_" & res
  End If
  'edit,delete and update
' Command2.Enabled = False
 'Command3.Enabled = False
' Command4.Enabled = False
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
If Text6.Text = "" Or Combo1.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
f = 1
ElseIf Not (DTPicker1.Value = Date) Then
MsgBox ("Enter today's date")
f = 1
ElseIf Val(Text6.Text) > Val(Text5.Text) Then
a = MsgBox("Quantity entered is more than pending quantity!", vbExclamation)
Text6.Text = ""
f = 1
ElseIf Val(Text6.Text) <= 0 Then
MsgBox ("Quantity cannot be less than 0")
Text6.Text = ""
f = 1
'to cal quantity to be deducted from purchase_stock
ElseIf z = 0 Then
 Set db1 = OpenDatabase("e:\\III year Project\PROJECT.MDB")
 Set rs1 = db1.OpenRecordset("select sur_name1, sur_name2 from sale_stock where sitem_no = '" & Text2.Text & "' ")
 rs1.MoveFirst
 '100
 n = rs1.Fields(0).Value
 'ml
 uom = rs1.Fields(1).Value
 'quantity produced
 q = Val(Text6.Text)
 Set rs1 = db1.OpenRecordset("select unit_of_measurement,stock_in_progress from purchase_stock where pitem_no = '" & Text4.Text & "' ")
 rs1.MoveFirst
 base_uom = rs1.Fields(0).Value
 sip = rs1.Fields(1).Value
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
 '
End If

If f = 0 Then
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Combo1.Text
Data1.Recordset.Fields(2) = DTPicker1.Value
Data1.Recordset.Fields(3) = Text2.Text
Data1.Recordset.Fields(4) = Text3.Text
Data1.Recordset.Fields(5) = Text4.Text
Data1.Recordset.Fields(6) = Val(Text6.Text)
'to reduce quantity, change status in production plan
Data2.Refresh
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
 If (Combo1.Text = Data2.Recordset.Fields(0)) Then
Data2.Recordset.Edit
Data2.Recordset.Fields(7) = Data2.Recordset.Fields(7) - q
 If Val(Text5.Text) = q Then
Data2.Recordset.Fields(6) = "Completed"
 End If
Data2.Recordset.Update
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop
'to add quantity in sale stock
Data3.Refresh
Data3.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
 If (Text2.Text = Data3.Recordset.Fields(0)) Then
Data3.Recordset.Edit
Data3.Recordset.Fields(7) = Data3.Recordset.Fields(7) + q
Data3.Recordset.Update
Exit Do
Else
Data3.Recordset.MoveNext
End If
Loop
'to reduce stock in progress in purchase stock
Data4.Refresh
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF
 If (Text4.Text = Data4.Recordset.Fields(0)) Then
Data4.Recordset.Edit
Data4.Recordset.Fields(5) = Data4.Recordset.Fields(5) - totalq
Data4.Recordset.Update
Exit Do
Else
Data4.Recordset.MoveNext
End If
Loop

Data1.Recordset.Update
MsgBox ("Production completion saved successfully")
'Navigation frame
Frame3.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
'edit,delete
'Command2.Enabled = True
'Command3.Enabled = True
'SELF
Command10.Enabled = False
 Command1.Enabled = True
 Command11.Enabled = True
End If

End Sub

Private Sub Command11_Click()
Me.Hide
Unload Me
Load Hom
Hom.Show
End Sub

Private Sub Command2_Click()
a = InputBox("Enter Production Complete Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
  If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
  Text1.Text = Data1.Recordset.Fields(0)
  DTPicker1.Value = Data1.Recordset.Fields(1)
  Combo1.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
  Text4.Text = Data1.Recordset.Fields(5)
  'add value to text5
  Set rs = db.OpenRecordset("select quantity from production_plan where pplan_no = '" & Combo1.Text & "' ")
  rs.MoveFirst
  Text5.Text = rs.Fields(0).Value

  Text6.Text = Data1.Recordset.Fields(6)
  DTPicker1.Enabled = False
  num = Data1.Recordset.Fields(5)
   flag = 1
  Exit Do
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
Else
MsgBox ("Record not found")
End If

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command9_Click()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Text5.Text = ""
 Text6.Text = ""
 DTPicker1.Value = Date
Combo1.ListIndex = -1
Command1.Enabled = True
 Command11.Enabled = True
End Sub

Private Sub Form_Activate()
'add items to combo
Combo1.Clear
Data2.Refresh
If Data2.Recordset.RecordCount <> 0 Then
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If Data2.Recordset.Fields(6) = "Pending" Then
Combo1.AddItem (Data2.Recordset.Fields(0))
Data2.Recordset.MoveNext
Else
Data2.Recordset.MoveNext
End If
Loop
Else
MsgBox ("No pendings in Production plan")
Unload Me
Load Hom
Hom.Show
End If

'edit,delete
Data1.Refresh
v = Data1.Recordset.RecordCount
If v = 0 Then
'Command2.Enabled = False
'Command3.Enabled = False
'navigation frame
Frame3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End If
Command10.Enabled = False
'Command4.Enabled = False
End Sub

Private Sub Text6_GotFocus()
'adding value to text
'sql
If Combo1.Text = "" Then
MsgBox ("Enter Production Plan Number first")
Else
Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
Set rs = db.OpenRecordset("select date,sitem_no,sitem_name, pitem_no,stock_in_progress from production_plan where pplan_no = '" & Combo1.Text & "' ")
rs.MoveFirst
DTPicker1.Value = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
End If

End Sub
