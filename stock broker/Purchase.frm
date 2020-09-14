VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Purchase 
   Caption         =   "k"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command14 
      Caption         =   "STOCK ARRIVAL ENTRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   48
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   18
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "supplier_stock"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "supplier"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      DataField       =   "area"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   45
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      DataField       =   "area"
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Top             =   6960
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3120
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "street"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   8
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "house_no"
      Height          =   285
      Left            =   3120
      TabIndex        =   7
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "supp_name"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   3360
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2880
      TabIndex        =   19
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H0080FFFF&
      TabIndex        =   20
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ORDER "
      Height          =   615
      Left            =   7440
      MaskColor       =   &H0080FFFF&
      TabIndex        =   16
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7200
      TabIndex        =   21
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7200
      TabIndex        =   22
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7200
      TabIndex        =   23
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   6960
      TabIndex        =   39
      Top             =   6120
      Width           =   2175
      Begin VB.CommandButton Command9 
         Caption         =   "CLEAR"
         Height          =   615
         Left            =   480
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   6960
      TabIndex        =   36
      Top             =   480
      Width           =   2175
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   960
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
         TabIndex        =   38
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   35
      Top             =   8400
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "PLACE NEW ORDER"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   24
      Top             =   480
      Width           =   6495
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   3015
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Calculate"
         Height          =   255
         Left            =   5280
         TabIndex        =   15
         Top             =   6960
         Width           =   975
      End
      Begin VB.CommandButton x 
         Caption         =   "Get data"
         Height          =   315
         Left            =   5280
         TabIndex        =   4
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cheque"
         Height          =   435
         Left            =   4560
         TabIndex        =   10
         Top             =   4200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cash"
         Height          =   495
         Left            =   3000
         TabIndex        =   9
         Top             =   4200
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   131137537
         CurrentDate     =   41058
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Purchase.frx":0000
         Left            =   3000
         List            =   "Purchase.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label12 
         Height          =   495
         Left            =   1440
         TabIndex        =   47
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Supplier Name"
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
         Left            =   480
         TabIndex        =   46
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Account Holder Name"
         Height          =   255
         Left            =   960
         TabIndex        =   44
         Top             =   5880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Account Number"
         Height          =   255
         Left            =   960
         TabIndex        =   43
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Cheque Number"
         Height          =   255
         Left            =   960
         TabIndex        =   42
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Purchase Item Name"
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
         Left            =   480
         TabIndex        =   41
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label16 
         Caption         =   "Purchase Order Number"
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
         Left            =   480
         TabIndex        =   34
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "Quantity"
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
         Left            =   480
         TabIndex        =   33
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label13 
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
         Left            =   480
         TabIndex        =   32
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Balance Amount"
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
         TabIndex        =   31
         Top             =   6960
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Amount Paid"
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
         TabIndex        =   30
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Total Amount"
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
         Left            =   480
         TabIndex        =   29
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Mode of Payment"
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
         Left            =   480
         TabIndex        =   28
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Supplier ID"
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
         Left            =   480
         TabIndex        =   27
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Purchase"
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
      Left            =   3360
      TabIndex        =   40
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mop As Integer
Public db As Database
Public rs As Recordset




Private Sub Combo2_Change()
'in case user changes purchase item number
Combo1.Enabled = False
Combo1.Clear
Text2.Text = ""
End Sub

Private Sub Command1_Click()
mop = 0
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
DTPicker1.Value = Date
Combo2.ListIndex = -1
Combo1.Clear
Option1.Value = False
Option2.Value = False
Combo1.Enabled = False
Command14.Enabled = False
'mode of payment
Label3.Visible = False
Label6.Visible = False
Label7.Visible = False
Text8.Visible = False
Text6.Visible = False
Text7.Visible = False
'add items to combo2
'sql
Data2.Refresh
Combo2.Clear
Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
Set rs = db.OpenRecordset("select distinct Pitem_no from supplier_stock")

If rs.EOF = True And rs.BOF = True Then
MsgBox ("Supplier stock table is empty. Cannot proceed to purchase")
Unload Me
Load Hom
Hom.Show

Exit Sub

Else
rs.MoveFirst
 Do While Not rs.EOF
 Combo2.AddItem (rs.Fields(0).Value)
 rs.MoveNext
 Loop
End If

'AUTO GENERATE
Data1.Refresh
v = Data1.Recordset.RecordCount
If v = 0 Then
  Text1.Text = "Porder_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 8)
  res = num + 1
  Text1.Text = "Porder_" & res
  End If
   'edit,delete and update
Command2.Enabled = False
' Command3.Enabled = False
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
z = 0
If Combo1.Text = "" Or Combo2.Text = "" Or Text4.Text = "" Or mop = 0 Or Text9.Text = "" Or Text10.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
z = 1
ElseIf Not (DTPicker1.Value = Date) Then
MsgBox ("Enter today's date")
z = 1
ElseIf Val(Text4.Text) <= 0 Then
MsgBox ("Quantity cannot be less than 0")
Text4.Text = ""
z = 1
ElseIf Val(Text9.Text) <= 0 Then
MsgBox ("Amount paid cannot be less than 0")
Text9.Text = ""
z = 1
ElseIf Option2.Value = True Then
 If Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
 b = MsgBox("Please fill all the contents!", vbExclamation)
 z = 1
 End If
End If
If z = 0 Then
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = DTPicker1.Value
Data1.Recordset.Fields(2) = Combo1.Text
Data1.Recordset.Fields(3) = Text3.Text
Data1.Recordset.Fields(4) = Combo2.Text
Data1.Recordset.Fields(5) = Text2.Text
'quantity arrived
Data1.Recordset.Fields(6) = 0
' quantity ordered
Data1.Recordset.Fields(15) = Val(Text4.Text)
Data1.Recordset.Fields(7) = Val(Text5.Text)

If Option1.Value = True Then
Data1.Recordset.Fields(8) = "cash"
Data1.Recordset.Fields(9) = "nil"
Data1.Recordset.Fields(10) = "nil"
Data1.Recordset.Fields(11) = "nil"
End If

If Option2.Value = True Then
Data1.Recordset.Fields(8) = "cheque"
Data1.Recordset.Fields(9) = Text6.Text
Data1.Recordset.Fields(10) = Text7.Text
Data1.Recordset.Fields(11) = Text8.Text
End If
Data1.Recordset.Fields(12) = Val(Text9.Text)
Data1.Recordset.Fields(13) = Val(Text10.Text)
Data1.Recordset.Fields(14) = "Pending"

Data1.Recordset.Update
MsgBox ("Order placed successfully")

'Navigation frame
Frame3.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
'edit,delete
Command2.Enabled = True
'Command3.Enabled = True
'SELF
Command10.Enabled = False

Command14.Enabled = True
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
'supplier name
If Combo1.Text = "" Then
MsgBox ("Please enter Supplier ID first")
Else
t = Combo1.Text
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If (t = Data2.Recordset.Fields(0)) Then
Text2.Text = Data2.Recordset.Fields(1)
Exit Do
Else
Data2.Recordset.MoveNext
 End If
Loop

'combo2
Combo2.Enabled = True
Combo2.Clear
z = Combo1.Text
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF
If (z = Data4.Recordset.Fields(0)) Then
Combo2.AddItem (Data4.Recordset.Fields(2))
End If
Data4.Recordset.MoveNext
Loop
End If
End Sub

Private Sub Command13_Click()
If Val(Text5.Text) < Val(Text9.Text) Then
MsgBox ("Paid amount is more than Total amount")
Text9.Text = ""
Text10.Text = ""
Else
Text10.Text = Val(Text5.Text) - Val(Text9.Text)
End If
End Sub

Private Sub Command14_Click()
On Error GoTo e

Stock_arrival.Show
Me.Hide
Exit Sub
e:
Resume Next
End Sub

Private Sub Command2_Click()
a = InputBox("Enter Purchase Order Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
 Text1.Text = Data1.Recordset.Fields(0)
 DTPicker1.Value = Data1.Recordset.Fields(1)
 ' add item to combo 2
 'Combo2.Clear
 'Combo2.AddItem (Data1.Recordset.Fields(4))
 'to add value to combo2
 'Combo2.Text = Data1.Recordset.Fields(4)

 Combo2.Clear
 Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
 Set rs = db.OpenRecordset("select distinct Pitem_no from supplier_stock")

If rs.EOF = True And rs.BOF = True Then
MsgBox ("Supplier stock table is empty. Cannot proceed to purchase")
Unload Me
Load Hom
Hom.Show
Exit Sub
Else
rs.MoveFirst
 Do While Not rs.EOF
 Combo2.AddItem (rs.Fields(0).Value)
 rs.MoveNext
 Loop
End If
Combo2.Text = Data1.Recordset.Fields(4)
 'add items to combo 1
 Combo1.Enabled = True
 Combo1.Clear
 Data4.Recordset.MoveFirst
 Do While Not Data4.Recordset.EOF
 If (Combo2.Text = Data4.Recordset.Fields(2)) Then
    Combo1.AddItem (Data4.Recordset.Fields(0))
    Data4.Recordset.MoveNext
 Else
 Data4.Recordset.MoveNext
 End If
 Loop

Combo1.Text = Data1.Recordset.Fields(2)
 'Combo1.AddItem (Data1.Recordset.Fields(2))
 'Combo1.Text = Data1.Recordset.Fields(2)
 
 Text3.Text = Data1.Recordset.Fields(3)
 'to add items to list in combo 1
 'to add unit of measurement in label 12
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
If (Combo2.Text = Data3.Recordset.Fields(0)) Then
Label12.Caption = "(in " + Data3.Recordset.Fields(3) + ")"
Exit Do
End If
Data3.Recordset.MoveNext
Loop

 Text2.Text = Data1.Recordset.Fields(5)
 Text4.Text = Data1.Recordset.Fields(15)
 Text5.Text = Data1.Recordset.Fields(7)
 If Data1.Recordset.Fields(8) = "cash" Then
 Option1.Value = True
 End If
 If Data1.Recordset.Fields(8) = "cheque" Then
 Option2.Value = True
Text6.Text = Data1.Recordset.Fields(9)
 Text7.Text = Data1.Recordset.Fields(10)
 Text8.Text = Data1.Recordset.Fields(11)
 End If

 Text9.Text = Data1.Recordset.Fields(12)
 Text10.Text = Data1.Recordset.Fields(13)
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
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
Command9.Enabled = False
Command10.Enabled = False
Command11.Enabled = False
Command14.Enabled = False
Else
MsgBox ("Record not found")
End If
End Sub

Private Sub Command3_Click()
a = InputBox("Enter Purchase Order Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
  Text1.Text = Data1.Recordset.Fields(0)
 DTPicker1.Value = Data1.Recordset.Fields(1)
 Combo1.Text = Data1.Recordset.Fields(2)
 Text2.Text = Data1.Recordset.Fields(3)
 Combo2.Enabled = True
Combo2.Clear
z = Combo1.Text
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF
If (z = Data4.Recordset.Fields(0)) Then
Combo2.AddItem (Data4.Recordset.Fields(2))
End If
Data4.Recordset.MoveNext
Loop
'to add value to combo2
Combo2.Text = Data1.Recordset.Fields(4)
'to add unit of measurement in label 12
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
If (Combo2.Text = Data3.Recordset.Fields(0)) Then
Label12.Caption = "(in " + Data3.Recordset.Fields(3) + ")"
Exit Do
End If
Data3.Recordset.MoveNext
Loop
 Text3.Text = Data1.Recordset.Fields(5)
 Text4.Text = Data1.Recordset.Fields(6)
 Text5.Text = Data1.Recordset.Fields(7)
 If Data1.Recordset.Fields(8) = "cash" Then
 Option1.Value = True
 End If
 If Data1.Recordset.Fields(8) = "cheque" Then
 Option2.Value = True
Text6.Text = Data1.Recordset.Fields(9)
 Text7.Text = Data1.Recordset.Fields(10)
 Text8.Text = Data1.Recordset.Fields(11)
 End If

 Text9.Text = Data1.Recordset.Fields(12)
 Text10.Text = Data1.Recordset.Fields(13)
 flag = 1
 Exit Do
 Else
 Data1.Recordset.MoveNext
 End If
Loop

If flag = 0 Then
MsgBox ("Record not found")
Else
z = Val(MsgBox("Do you want to delete this record?", vbYesNo))
  If (z = 6) Then
  Data1.Recordset.Delete
  MsgBox ("Record deleted successfully")
  End If
   Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
 Text8.Text = ""
 Text9.Text = ""
 Text10.Text = ""
 Combo1.ListIndex = -1
 Combo2.ListIndex = -1
 Option1.Value = False
 Option2.Value = False
 Combo2.Enabled = False
End If

'NAvigation frame
v = Data1.Recordset.RecordCount
If v = 0 Then
Frame3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Command7.Enabled = False
Command8.Enabled = False
End If
End Sub

Private Sub Command4_Click()
z = 0
If Combo1.Text = "" Or Combo2.Text = "" Or Text4.Text = "" Or mop = 0 Or Text9.Text = "" Or Text10.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
z = 1
ElseIf Val(Text4.Text) <= 0 Then
MsgBox ("Quantity cannot be less than 0")
Text4.Text = ""
z = 1
ElseIf Val(Text9.Text) <= 0 Then
MsgBox ("Amount paid cannot be less than 0")
Text9.Text = ""
z = 1
ElseIf Option2.Value = True Then
 If Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
 b = MsgBox("Please fill all the contents!", vbExclamation)
 z = 1
 End If
End If
If z = 0 Then
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = DTPicker1.Value
Data1.Recordset.Fields(2) = Combo1.Text
Data1.Recordset.Fields(3) = Text3.Text
Data1.Recordset.Fields(4) = Combo2.Text
Data1.Recordset.Fields(5) = Text2.Text
Data1.Recordset.Fields(15) = Val(Text4.Text)
Data1.Recordset.Fields(7) = Val(Text5.Text)

If Option1.Value = True Then
Data1.Recordset.Fields(8) = "cash"
Data1.Recordset.Fields(9) = "nil"
Data1.Recordset.Fields(10) = "nil"
Data1.Recordset.Fields(11) = "nil"
End If

If Option2.Value = True Then
Data1.Recordset.Fields(8) = "cheque"
Data1.Recordset.Fields(9) = Text6.Text
Data1.Recordset.Fields(10) = Text7.Text
Data1.Recordset.Fields(11) = Text8.Text
End If
Data1.Recordset.Fields(12) = Val(Text9.Text)
Data1.Recordset.Fields(13) = Val(Text10.Text)
Data1.Recordset.Update
MsgBox ("Updated successfully")
Frame3.Enabled = True
Frame4.Enabled = True
Command1.Enabled = True
'Command3.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Enabled = True
Command11.Enabled = True
Command14.Enabled = True
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
 Combo1.AddItem (Data1.Recordset.Fields(2))
 Combo1.Text = Data1.Recordset.Fields(2)
 Text3.Text = Data1.Recordset.Fields(3)
 Combo1.Enabled = True
 Combo2.AddItem (Data1.Recordset.Fields(4))
 Combo2.Text = Data1.Recordset.Fields(4)
 Text2.Text = Data1.Recordset.Fields(5)
 Text4.Text = Data1.Recordset.Fields(15)
 Text5.Text = Data1.Recordset.Fields(7)
 If Data1.Recordset.Fields(8) = "cash" Then
 Option1.Value = True
 End If
 If Data1.Recordset.Fields(8) = "cheque" Then
 Option2.Value = True
 Text6.Text = Data1.Recordset.Fields(9)
 Text7.Text = Data1.Recordset.Fields(10)
 Text8.Text = Data1.Recordset.Fields(11)
 End If
 Text9.Text = Data1.Recordset.Fields(12)
 Text10.Text = Data1.Recordset.Fields(13)

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
 Combo1.AddItem (Data1.Recordset.Fields(2))
 Combo1.Text = Data1.Recordset.Fields(2)
 Text3.Text = Data1.Recordset.Fields(3)
 Combo1.Enabled = True
 Combo2.AddItem (Data1.Recordset.Fields(4))
 Combo2.Text = Data1.Recordset.Fields(4)
 Text2.Text = Data1.Recordset.Fields(5)
 Text4.Text = Data1.Recordset.Fields(15)
 Text5.Text = Data1.Recordset.Fields(7)
 If Data1.Recordset.Fields(8) = "cash" Then
 Option1.Value = True
 End If
 If Data1.Recordset.Fields(8) = "cheque" Then
 Option2.Value = True
 Text6.Text = Data1.Recordset.Fields(9)
 Text7.Text = Data1.Recordset.Fields(10)
 Text8.Text = Data1.Recordset.Fields(11)
 End If
 Text9.Text = Data1.Recordset.Fields(12)
 Text10.Text = Data1.Recordset.Fields(13)
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
 Combo1.AddItem (Data1.Recordset.Fields(2))
 Combo1.Text = Data1.Recordset.Fields(2)
 Text3.Text = Data1.Recordset.Fields(3)
 Combo1.Enabled = True
 Combo2.AddItem (Data1.Recordset.Fields(4))
 Combo2.Text = Data1.Recordset.Fields(4)
 Text2.Text = Data1.Recordset.Fields(5)
 Text4.Text = Data1.Recordset.Fields(15)
 Text5.Text = Data1.Recordset.Fields(7)
 If Data1.Recordset.Fields(8) = "cash" Then
 Option1.Value = True
 End If
 If Data1.Recordset.Fields(8) = "cheque" Then
 Option2.Value = True
 Text6.Text = Data1.Recordset.Fields(9)
 Text7.Text = Data1.Recordset.Fields(10)
 Text8.Text = Data1.Recordset.Fields(11)
 End If
 Text9.Text = Data1.Recordset.Fields(12)
 Text10.Text = Data1.Recordset.Fields(13)
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
 Combo1.AddItem (Data1.Recordset.Fields(2))
 Combo1.Text = Data1.Recordset.Fields(2)
 Text3.Text = Data1.Recordset.Fields(3)
 Combo1.Enabled = True
 Combo2.AddItem (Data1.Recordset.Fields(4))
 Combo2.Text = Data1.Recordset.Fields(4)
 Text2.Text = Data1.Recordset.Fields(5)
 Text4.Text = Data1.Recordset.Fields(15)
 Text5.Text = Data1.Recordset.Fields(7)
 If Data1.Recordset.Fields(8) = "cash" Then
 Option1.Value = True
 End If
 If Data1.Recordset.Fields(8) = "cheque" Then
 Option2.Value = True
 Text6.Text = Data1.Recordset.Fields(9)
 Text7.Text = Data1.Recordset.Fields(10)
 Text8.Text = Data1.Recordset.Fields(11)
 End If
 Text9.Text = Data1.Recordset.Fields(12)
 Text10.Text = Data1.Recordset.Fields(13)
 Command7.Enabled = False
 Command6.Enabled = True
 End If

End Sub

Private Sub Command9_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
DTPicker1.Value = Date
Combo2.ListIndex = -1
Combo1.Clear
Option1.Value = False
Option2.Value = False
Combo1.Enabled = False
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
'save
Command10.Enabled = False
Command1.Enabled = True
 Command11.Enabled = True
Command14.Enabled = True
End Sub

Private Sub Form_Activate()
DTPicker1.Value = Date
'add items to combo2
'sql
Data2.Refresh
Combo2.Clear
Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
Set rs = db.OpenRecordset("select distinct Pitem_no from supplier_stock")

If rs.EOF = True And rs.BOF = True Then
MsgBox ("Supplier stock table is empty. Cannot proceed to purchase")
Unload Me
Load Hom
Hom.Show

Exit Sub

Else
rs.MoveFirst
 Do While Not rs.EOF
 Combo2.AddItem (rs.Fields(0).Value)
 rs.MoveNext
 Loop
End If

'combo1
Combo1.Enabled = False
'disable
Data1.Refresh
v = Data1.Recordset.RecordCount
'edit,delete
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
'Save
Command10.Enabled = False
Command4.Enabled = False
End Sub


Private Sub Option1_Click()
mop = 1
Label3.Visible = False
Label6.Visible = False
Label7.Visible = False
Text8.Visible = False
Text6.Visible = False
Text7.Visible = False
End Sub

Private Sub Option2_Click()
mop = 2
Label3.Visible = True
Label6.Visible = True
Label7.Visible = True
Text8.Visible = True
Text6.Visible = True
Text7.Visible = True
End Sub

Private Sub Text2_Change()
If IsNumeric(Text2.Text) Then
Text2.Text = ""
End If
End Sub

Private Sub Text3_Change()
If IsNumeric(Text3.Text) Then
Text3.Text = ""
End If

End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4.Text) Then
Text4.Text = ""
End If

End Sub

Private Sub Text4_GotFocus()
'supplier name
If Combo1.Text = "" Then
MsgBox (" Enter Supplier ID")
Else
t = Combo1.Text
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If (t = Data2.Recordset.Fields(0)) Then
Text3.Text = Data2.Recordset.Fields(1)
Exit Do
Else
Data2.Recordset.MoveNext
 End If
Loop

'to add unit of measurement in label 12
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
If (Combo2.Text = Data3.Recordset.Fields(0)) Then
Label12.Caption = "(in " + Data3.Recordset.Fields(3) + ")"
Exit Do
End If
Data3.Recordset.MoveNext
Loop
End If
End Sub

Private Sub Text4_LostFocus()
'amount
q = Val(Text4.Text)
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF
If (Combo1.Text = Data4.Recordset.Fields(0) And Combo2.Text = Data4.Recordset.Fields(2)) Then
amt = q * Val(Data4.Recordset.Fields(4))
Text5.Text = amt
Exit Do
Else
Data4.Recordset.MoveNext
End If
Loop
End Sub

Private Sub Text6_Change()
If Not IsNumeric(Text6.Text) Then
Text6.Text = ""
End If

End Sub

Private Sub Text7_Change()
If Not IsNumeric(Text7.Text) Then
Text7.Text = ""
End If

End Sub

Private Sub Text8_Change()
If IsNumeric(Text8.Text) Then
Text8.Text = ""
End If

End Sub

Private Sub Text8_LostFocus()
amt = Val(Text4.Text)
paid = Val(Text8.Text)
Text9.Text = amt - paid
End Sub

Private Sub Text9_Change()
If Not IsNumeric(Text9.Text) Then
Text9.Text = ""
End If

Text10.Text = ""
End Sub

Private Sub x_Click()
'purchase item name
If Combo2.Text = "" Then
MsgBox ("Please enter Purchase Item Number first")
Else
t = Combo2.Text
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
If (t = Data3.Recordset.Fields(0)) Then
Text2.Text = Data3.Recordset.Fields(1)
Exit Do
Else
Data3.Recordset.MoveNext
 End If
Loop

'combo1
Combo1.Enabled = True
Combo1.Clear
z = Combo2.Text
Data4.Recordset.MoveFirst
Do While Not Data4.Recordset.EOF
If (z = Data4.Recordset.Fields(2)) Then
    Combo1.AddItem (Data4.Recordset.Fields(0))
    Data4.Recordset.MoveNext
Else
Data4.Recordset.MoveNext
End If
Loop
End If
End Sub
