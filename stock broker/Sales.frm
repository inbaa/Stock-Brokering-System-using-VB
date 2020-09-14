VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Sales 
   Caption         =   "Sales"
   ClientHeight    =   9645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   9645
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7320
      MaskColor       =   &H0080FFFF&
      TabIndex        =   11
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "ORDER "
      Height          =   615
      Left            =   7320
      MaskColor       =   &H0080FFFF&
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      DataField       =   "area"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Top             =   7440
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      DataField       =   "area"
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Top             =   6960
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3000
      TabIndex        =   4
      Top             =   5880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   5400
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "street"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "house_no"
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   3840
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "supp_name"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7080
      TabIndex        =   13
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7080
      TabIndex        =   14
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   6495
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Sales.frx":0000
         Left            =   2880
         List            =   "Sales.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1440
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cash"
         Height          =   495
         Left            =   3000
         TabIndex        =   27
         Top             =   4200
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cheque"
         Height          =   435
         Left            =   4560
         TabIndex        =   26
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Calculate"
         Height          =   255
         Left            =   5280
         TabIndex        =   25
         Top             =   6960
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2400
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Top             =   840
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Format          =   136708097
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
         TabIndex        =   45
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Left            =   480
         TabIndex        =   44
         Top             =   1440
         Width           =   1815
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
         TabIndex        =   43
         Top             =   4320
         Width           =   1695
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
         TabIndex        =   42
         Top             =   3840
         Width           =   1455
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
         TabIndex        =   41
         Top             =   6480
         Width           =   1575
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
         TabIndex        =   40
         Top             =   6960
         Width           =   1815
      End
      Begin VB.Label Label13 
         Caption         =   "Customer ID"
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
         TabIndex        =   39
         Top             =   2400
         Width           =   1935
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
         TabIndex        =   38
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "Sale Number"
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
         TabIndex        =   37
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Customer Name"
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
         TabIndex        =   36
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Cheque Number"
         Height          =   255
         Left            =   960
         TabIndex        =   35
         Top             =   4920
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Account Number"
         Height          =   255
         Left            =   960
         TabIndex        =   34
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Account Holder Name"
         Height          =   255
         Left            =   960
         TabIndex        =   33
         Top             =   5880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label11 
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
         Left            =   480
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label12 
         Height          =   495
         Left            =   1440
         TabIndex        =   31
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   8400
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "PLACE NEW ORDER"
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4215
      Left            =   6840
      TabIndex        =   18
      Top             =   480
      Width           =   2175
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   19
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
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   6840
      TabIndex        =   16
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sales"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sale_stock"
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
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "customer"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "supplier_stock"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label21 
      Caption         =   "Sales"
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
      Left            =   3480
      TabIndex        =   46
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mop As Integer
Public db As Database
Public rs As Recordset
Private Sub Combo1_Change()
Text2.Text = ""
End Sub

Private Sub Combo2_Change()
Text3.Text = ""
End Sub

Private Sub Combo2_Click()
'sale item name
t = Combo1.Text
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If (t = Data2.Recordset.Fields(0)) Then
Text2.Text = Data2.Recordset.Fields(5)
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop

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
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Option1.Value = False
Option2.Value = False

'mode of payment
Label3.Visible = False
Label6.Visible = False
Label7.Visible = False
Text8.Visible = False
Text6.Visible = False
Text7.Visible = False
'AUTO GENERATE
Data1.Refresh
v = Data1.Recordset.RecordCount
If v = 0 Then
  Text1.Text = "Sale_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 6)
  res = num + 1
  Text1.Text = "Sale_" & res
  End If
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
Data1.Recordset.Fields(3) = Text2.Text
Data1.Recordset.Fields(4) = Combo2.Text
Data1.Recordset.Fields(5) = Text3.Text
Data1.Recordset.Fields(6) = Val(Text4.Text)
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
'to subract quantity from sale stock
Data2.Refresh
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
 If (Combo1.Text = Data2.Recordset.Fields(0)) Then
 ' to check if quantity is available
  If (Data2.Recordset.Fields(7) - Val(Text4.Text) > 0) Then
 Data2.Recordset.Edit
 Data2.Recordset.Fields(7) = Data2.Recordset.Fields(7) - Val(Text4.Text)
 Data2.Recordset.Update
 z = 2
  Else
 b = MsgBox("Quantity is not available!", vbExclamation)
 Text4.Text = ""
 Text5.Text = ""
 Text9.Text = ""
 Text10.Text = ""
  End If
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop
 If z = 2 Then
Data1.Recordset.Update
MsgBox ("Sold successfully")
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
Command1.Enabled = True
 Command11.Enabled = True
 End If
End If
End Sub

Private Sub Command11_Click()
Me.Hide
Unload Me
Hom.Show
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

Private Sub Command2_Click()
a = InputBox("Enter Purchase Order Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
 Text1.Text = Data1.Recordset.Fields(0)
 DTPicker1.Value = Data1.Recordset.Fields(1)
 Combo1.Text = Data1.Recordset.Fields(2)
 Text2.Text = Data1.Recordset.Fields(3)
 Combo2.Text = Data1.Recordset.Fields(4)
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
Else
MsgBox ("Record not found")
End If

End Sub

Private Sub Command3_Click()

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
Data1.Recordset.Fields(3) = Text2.Text
Data1.Recordset.Fields(4) = Combo2.Text
Data1.Recordset.Fields(5) = Text3.Text

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
' to check for change in quantity
If Data1.Recordset.Fields(6) < Val(Text4.Text) Then
fq = Val(Text4.Text) - Data1.Recordset.Fields(6)
' subract
x = 1
ElseIf Data1.Recordset.Fields(6) > Val(Text4.Text) Then
fq = Data1.Recordset.Fields(6) - Val(Text4.Text)
' add
x = 2
ElseIf Data1.Recordset.Fields(6) = Val(Text4.Text) Then
Data1.Recordset.Fields(6) = Val(Text4.Text)
End If
' to check if quantity is available
 
 Data2.Recordset.Edit
 If x = 1 Then
  If (Data2.Recordset.Fields(7) - fq > 0) Then
  Data2.Recordset.Fields(7) = Data2.Recordset.Fields(7) - fq
  Data2.Recordset.Update
  z = 4
  Else
  z = 3
  End If
 ElseIf x = 2 Then
  Data2.Recordset.Fields(7) = Data2.Recordset.Fields(7) + fq
   z = 4
   Data2.Recordset.Update
 End If
If z = 3 Then
b = MsgBox("Quantity is not available!", vbExclamation)
ElseIf z = 4 Then
Data1.Recordset.Fields(6) = Val(Text4.Text)
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
'SELF
Command4.Enabled = False
End If
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
Combo1.ListIndex = -1
Combo2.Clear
Option1.Value = False
Option2.Value = False
Combo2.Enabled = False
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

End Sub

Private Sub Form_Activate()
DTPicker1.Value = Date
'add items to combo
Combo1.Clear
If Data2.Recordset.RecordCount <> 0 Then
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Combo1.AddItem (Data2.Recordset.Fields(0))
Data2.Recordset.MoveNext
Loop
Else
MsgBox ("Sale stock table is empty")
Me.Hide
Unload Me
Hom.Show
End If
'add items to combo2
Combo2.Clear
If Data3.Recordset.RecordCount <> 0 Then
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
Combo2.AddItem (Data3.Recordset.Fields(0))
Data3.Recordset.MoveNext
Loop
Else
MsgBox ("Customer table is empty")
Me.Hide
Unload Me
Hom.Show
End If

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

Private Sub Text4_Change()
Text9.Text = ""
Text10.Text = ""
End Sub

Private Sub Text4_GotFocus()
'sale item name
t = Combo2.Text
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
If (t = Data3.Recordset.Fields(0)) Then
Text3.Text = Data3.Recordset.Fields(1)
Exit Do
Else
Data3.Recordset.MoveNext
End If
Loop

End Sub

Private Sub Text4_LostFocus()
'amount
q = Val(Text4.Text)
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If (Combo1.Text = Data2.Recordset.Fields(0)) Then
amt = q * Val(Data2.Recordset.Fields(8))
Text5.Text = amt
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop
End Sub

Private Sub Text9_Change()
If Not IsNumeric(Text9.Text) Then
Text9.Text = ""
End If

Text10.Text = ""

End Sub
