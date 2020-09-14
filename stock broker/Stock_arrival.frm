VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Stock_arrival 
   Caption         =   "Stock Arrival"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   4080
      TabIndex        =   24
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "NEW STOCK ARRIVAL ENTRY"
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      DataField       =   "street"
      Height          =   285
      Left            =   3360
      TabIndex        =   22
      Top             =   4800
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      DataField       =   "street"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   21
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "street"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   1
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "house_no"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "supp_name"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   6495
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Stock_arrival.frx":0000
         Left            =   3000
         List            =   "Stock_arrival.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   132120577
         CurrentDate     =   41058
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
         TabIndex        =   19
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         TabIndex        =   18
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Quantity Ordered"
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
         TabIndex        =   17
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Quantity Arrived"
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
         TabIndex        =   16
         Top             =   4320
         Width           =   1815
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
         TabIndex        =   15
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label14 
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
         TabIndex        =   14
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Stock Arrival Number"
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
         TabIndex        =   13
         Top             =   360
         Width           =   2055
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
         TabIndex        =   12
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label12 
         Height          =   495
         Left            =   1440
         TabIndex        =   10
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "stock_arrival"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   2400
      MaskColor       =   &H0080FFFF&
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   615
      Left            =   5880
      MaskColor       =   &H0080FFFF&
      TabIndex        =   4
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label21 
      Caption         =   "Stock Arrival Entry"
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
      TabIndex        =   20
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "Stock_arrival"
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
Text7.Text = ""
DTPicker1.Value = Date
Combo1.ListIndex = -1
'add items to combo1
'sql
Data2.Refresh
Combo1.Clear
Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
Set rs = db.OpenRecordset("select Porder_no from purchase where status='Pending' ")

If rs.EOF = True Or rs.BOF = True Then
MsgBox ("Purchase order table is empty. Cannot proceed to stock arrival")
Hom.Show
Me.Hide
Unload Me
Exit Sub
Else
rs.MoveFirst
 Do While Not rs.EOF
 Combo1.AddItem (rs.Fields(0).Value)
 rs.MoveNext
 Loop
End If

'AUTO GENERATE
Data1.Refresh
v = Data1.Recordset.RecordCount
If v = 0 Then
  Text1.Text = "Stock_arrival_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 15)
  res = num + 1
  Text1.Text = "Stock_arrival_" & res
  End If
Command2.Enabled = True

Command1.Enabled = False
 Command3.Enabled = False
End Sub

Private Sub Command9_Click()

End Sub

Private Sub Command2_Click()
z = 0
If Combo1.Text = "" Or Text7.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
z = 1
ElseIf Not (DTPicker1.Value = Date) Then
MsgBox ("Enter today's date")
z = 1
ElseIf Val(Text7.Text) > Val(Text6.Text) Then
MsgBox ("Quantity arrived is more than Quantity ordered")
Text7.Text = ""
z = 1
ElseIf Val(Text7.Text) <= 0 Then
MsgBox ("Quantity arrived cannot be less than 0")
Text7.Text = ""
z = 1
End If

If z = 0 Then
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = DTPicker1.Value
Data1.Recordset.Fields(2) = Combo1.Text
Data1.Recordset.Fields(3) = Val(Text7.Text)

'to update quantity in purchase stock
Data3.Refresh
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
 If (Text4.Text = Data3.Recordset.Fields(0)) Then
Data3.Recordset.Edit
Data3.Recordset.Fields(2) = Data3.Recordset.Fields(2) + Val(Text7.Text)
Data3.Recordset.Update
Exit Do
Else
Data3.Recordset.MoveNext
End If
Loop
'purchase status and quantity
Data2.Refresh
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
 If (Combo1.Text = Data2.Recordset.Fields(0)) Then
Data2.Recordset.Edit
Data2.Recordset.Fields(6) = Data2.Recordset.Fields(6) + Val(Text7.Text)
 If Val(Text7.Text) = Val(Text6.Text) Then
Data2.Recordset.Fields(14) = "Arrived"
 End If
Data2.Recordset.Update
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop

Data1.Recordset.Update
MsgBox ("Data saved successfully")
'SELF
Command2.Enabled = False
Command1.Enabled = True
 Command3.Enabled = True
End If

End Sub

Private Sub Command3_Click()
Me.Hide
Unload Me
Load Hom
Hom.Show

End Sub

Private Sub Command4_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
DTPicker1.Value = Date
Combo1.ListIndex = -1
Command2.Enabled = False
Command1.Enabled = True
 Command3.Enabled = True
End Sub

Private Sub Form_Load()
DTPicker1.Value = Date
'add items to combo1
'sql
Data2.Refresh
Combo1.Clear
Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
Set rs = db.OpenRecordset("select Porder_no from purchase where status='Pending' ")

If rs.EOF = True And rs.BOF = True Then
MsgBox ("Purchase order table is empty. Cannot proceed to stock arrival")
Unload Me
Load Hom
Hom.Show

Exit Sub

Else
rs.MoveFirst
 Do While Not rs.EOF
 Combo1.AddItem (rs.Fields(0).Value)
 rs.MoveNext
 Loop
End If

Command2.Enabled = False

End Sub

Private Sub Frame4_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label9_Click()
End Sub

Private Sub Text7_Change()
If Not IsNumeric(Text7.Text) Then
Text7.Text = ""
End If

End Sub

Private Sub Text7_GotFocus()
'adding value to text
'sql
If Combo1.Text = "" Then
MsgBox ("Enter Purchase Order Number first")
Else

Set db = OpenDatabase("e:\\III year Project\PROJECT.MDB")
Set rs = db.OpenRecordset("select supp_id,supp_name,pitem_no,pitem_name,quantity_ordered from purchase where porder_no = '" & Combo1.Text & "' ")
rs.MoveFirst
Text2.Text = rs.Fields(0).Value
Text3.Text = rs.Fields(1).Value
Text4.Text = rs.Fields(2).Value
Text5.Text = rs.Fields(3).Value
Text6.Text = rs.Fields(4).Value
End If

End Sub
