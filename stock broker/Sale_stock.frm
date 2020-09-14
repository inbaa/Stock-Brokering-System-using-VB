VERSION 5.00
Begin VB.Form Sale_stock 
   Caption         =   "Sale_stock"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command15 
      Caption         =   "30%"
      Height          =   375
      Left            =   4920
      TabIndex        =   42
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      Caption         =   "20%"
      Height          =   375
      Left            =   3480
      TabIndex        =   41
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      DataField       =   "street"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   37
      Top             =   4680
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      DataField       =   "street"
      Height          =   285
      Left            =   3000
      TabIndex        =   36
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      DataField       =   "street"
      Height          =   285
      Left            =   3000
      TabIndex        =   34
      Top             =   3960
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      DataField       =   "street"
      Height          =   285
      Left            =   3000
      TabIndex        =   33
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "street"
      Height          =   285
      Left            =   3000
      TabIndex        =   8
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2760
      TabIndex        =   12
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataField       =   "street"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   7
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "house_no"
      Height          =   285
      Left            =   3840
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "supp_name"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   3015
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H0080FFFF&
      TabIndex        =   13
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H0080FFFF&
      TabIndex        =   9
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7320
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7320
      TabIndex        =   15
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7320
      TabIndex        =   16
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "sale_stock"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   7080
      TabIndex        =   29
      Top             =   4320
      Width           =   2295
      Begin VB.CommandButton Command9 
         Caption         =   "CLEAR"
         Height          =   615
         Left            =   360
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   7080
      TabIndex        =   26
      Top             =   480
      Width           =   2295
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   27
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
         TabIndex        =   28
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000017&
      Height          =   1095
      Left            =   240
      TabIndex        =   25
      Top             =   6000
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "ADD NEW"
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5175
      Left            =   120
      TabIndex        =   17
      Top             =   480
      Width           =   6855
      Begin VB.CommandButton Command13 
         Caption         =   "10%"
         Height          =   375
         Left            =   2040
         TabIndex        =   40
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Get data"
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   18
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Sale_stock.frx":0000
         Left            =   2760
         List            =   "Sale_stock.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label11 
         Caption         =   "Profit"
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
         TabIndex        =   43
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label10 
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
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Transportation Charge"
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
         TabIndex        =   38
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Packing Charge"
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
         TabIndex        =   35
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Rate"
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
         TabIndex        =   32
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "To be prdouced in:-"
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
         TabIndex        =   31
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label18 
         Caption         =   "@"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label Label1 
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
         TabIndex        =   23
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label5 
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
         Left            =   360
         TabIndex        =   22
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Minimum Level (in quantity)"
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
         Top             =   2760
         Width           =   2415
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
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label20 
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
         TabIndex        =   19
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
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
      Left            =   3960
      TabIndex        =   30
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "Sales Stock"
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
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Sale_stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p As Variant
Private Sub Combo1_Click()
'in case user changes selected pitem_no
Combo2.Enabled = False
Combo2.Clear
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Combo2_LostFocus()
Text4.Text = Text2.Text + " " + Text3.Text + " " + Combo2.Text

End Sub

Private Sub Command1_Click()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Text5.Text = ""
 Text6.Text = ""
 Text7.Text = ""
 Text8.Text = ""
 Text9.Text = ""
 Combo1.ListIndex = -1
 Combo2.Clear
'AUTO GENERATE
 Data1.Refresh
 v = Data1.Recordset.RecordCount
 If v = 0 Then
  Text1.Text = "Sitem_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 7)
  res = num + 1
  Text1.Text = "Sitem_" & res
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
f = 0
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
f = 1
ElseIf Val(Text5.Text) <= 0 Then
MsgBox ("Minimum Quantity cannot be less than 0")
Text5.Text = ""
f = 1
ElseIf Val(Text6.Text) <= 0 Or Val(Text7.Text) <= 0 Or Val(Text8.Text) <= 0 Then
MsgBox ("Amount cannot be less than 0")
f = 1
'to check for duplicate
ElseIf f = 0 Then
 sitem = Text4.Text
 Do While Not Data1.Recordset.EOF
 If sitem = Data1.Recordset.Fields(5) Then
 b = MsgBox("Given Sale item already entered!", vbExclamation)
 f = 1
 Exit Do
 Else
 Data1.Recordset.MoveNext
 End If
 Loop
End If
If f = 0 Then
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Combo1.Text
Data1.Recordset.Fields(2) = Text2.Text
Data1.Recordset.Fields(3) = Val(Text3.Text)
Data1.Recordset.Fields(4) = Combo2.Text
Data1.Recordset.Fields(5) = Text4.Text
Data1.Recordset.Fields(6) = Val(Text5.Text)
Data1.Recordset.Fields(7) = 0
Data1.Recordset.Fields(8) = Val(Text9.Text)
Data1.Recordset.Fields(9) = Val(Text7.Text)
Data1.Recordset.Fields(10) = Val(Text8.Text)
Data1.Recordset.Fields(11) = Val(Text6.Text)
If p = 10 Then
Data1.Recordset.Fields(12) = "10%"
ElseIf p = 20 Then
Data1.Recordset.Fields(12) = "20%"
ElseIf p = 30 Then
Data1.Recordset.Fields(12) = "30%"
End If
Data1.Recordset.Update
MsgBox ("Sale stock details saved successfully")
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

End Sub

Private Sub Command11_Click()
Me.Hide
Unload Me
Hom.Show
End Sub

Private Sub Command12_Click()
'purchase item name
If Combo1.Text = "" Then
MsgBox ("Please enter Purchase Item name first")
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
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If (z = Data2.Recordset.Fields(0)) Then
   If (Data2.Recordset.Fields(3) = "litres") Then
   Combo2.AddItem ("ml")
   Combo2.AddItem ("litres")
   ElseIf (Data2.Recordset.Fields(3) = "kilogram" Or Data2.Recordset.Fields(3) = "ton") Then
   Combo2.AddItem ("mg")
   Combo2.AddItem ("kg")
   End If
   Exit Do
End If
Data2.Recordset.MoveNext
Loop
End If
End Sub

Private Sub Command13_Click()
Text9.Text = ""
If Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
a = MsgBox("Please fill all rates!", vbExclamation)
Else
x = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
x = x + x * 0.1
Text9.Text = x
p = 10
End If
End Sub

Private Sub Command14_Click()
Text9.Text = ""
If Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
a = MsgBox("Please fill all rates!", vbExclamation)
Else
x = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
x = x + x * 0.2
Text9.Text = x
p = 20
End If
End Sub

Private Sub Command15_Click()
Text9.Text = ""
If Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
a = MsgBox("Please fill all rates!", vbExclamation)
Else
x = Val(Text6.Text) + Val(Text7.Text) + Val(Text8.Text)
x = x + x * 0.3
Text9.Text = x
p = 30
End If
End Sub

Private Sub Command2_Click()
a = InputBox("Enter Sale Item Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
Text1.Text = Data1.Recordset.Fields(0)
Combo1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)
Text3.Text = Data1.Recordset.Fields(3)
Combo2.AddItem (Data1.Recordset.Fields(4))
Combo2.Text = Data1.Recordset.Fields(4)
Text4.Text = Data1.Recordset.Fields(5)
Text5.Text = Data1.Recordset.Fields(6)
Text9.Text = Data1.Recordset.Fields(8)
Text7.Text = Data1.Recordset.Fields(9)
Text8.Text = Data1.Recordset.Fields(10)
Text6.Text = Data1.Recordset.Fields(11)
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
a = InputBox("Enter Sale Item Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
Text1.Text = Data1.Recordset.Fields(0)
Combo1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)
Text3.Text = Data1.Recordset.Fields(3)
Combo2.AddItem (Data1.Recordset.Fields(4))
Combo2.Text = Data1.Recordset.Fields(4)
Text4.Text = Data1.Recordset.Fields(5)
Text5.Text = Data1.Recordset.Fields(6)
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
 Combo1.ListIndex = -1
Combo2.Clear
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
f = 0
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
f = 1
ElseIf Val(Text5.Text) <= 0 Then
MsgBox ("Minimum Quantity cannot be less than 0")
Text5.Text = ""
f = 1
ElseIf Val(Text6.Text) <= 0 Or Val(Text7.Text) <= 0 Or Val(Text8.Text) <= 0 Then
MsgBox ("Amount cannot be less than 0")
f = 1
'to check for duplicate
ElseIf f = 0 Then
 sitem = Text4.Text
 Do While Not Data1.Recordset.EOF
 If sitem = Data1.Recordset.Fields(5) Then
 b = MsgBox("Given Sale item already entered!", vbExclamation)
 f = 1
 Exit Do
 Else
 Data1.Recordset.MoveNext
 End If
 Loop
End If
If f = 0 Then
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Combo1.Text
Data1.Recordset.Fields(2) = Text2.Text
Data1.Recordset.Fields(3) = Val(Text3.Text)
Data1.Recordset.Fields(4) = Combo2.Text
Data1.Recordset.Fields(5) = Text4.Text
Data1.Recordset.Fields(6) = Val(Text5.Text)
Data1.Recordset.Fields(8) = Val(Text9.Text)
Data1.Recordset.Fields(9) = Val(Text7.Text)
Data1.Recordset.Fields(10) = Val(Text8.Text)
Data1.Recordset.Fields(11) = Val(Text6.Text)
If p = 10 Then
Data1.Recordset.Fields(12) = "10%"
ElseIf p = 20 Then
Data1.Recordset.Fields(12) = "20%"
ElseIf p = 30 Then
Data1.Recordset.Fields(12) = "30%"
End If
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
End Sub

Private Sub Command5_Click()
Data1.Recordset.MoveFirst
 If Data1.Recordset.BOF = True Then
MsgBox ("No record to display")
Command5.Enabled = False
Else
Text1.Text = Data1.Recordset.Fields(0)
Combo1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)
Text3.Text = Data1.Recordset.Fields(3)
Combo2.AddItem (Data1.Recordset.Fields(4))
Combo2.Text = Data1.Recordset.Fields(4)
Text4.Text = Data1.Recordset.Fields(5)
Text5.Text = Data1.Recordset.Fields(6)
Text9.Text = Data1.Recordset.Fields(8)
Text7.Text = Data1.Recordset.Fields(9)
Text8.Text = Data1.Recordset.Fields(10)
Text6.Text = Data1.Recordset.Fields(11)
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
Combo1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)
Text3.Text = Data1.Recordset.Fields(3)
Combo2.AddItem (Data1.Recordset.Fields(4))
Combo2.Text = Data1.Recordset.Fields(4)
Text4.Text = Data1.Recordset.Fields(5)
Text5.Text = Data1.Recordset.Fields(6)
Text9.Text = Data1.Recordset.Fields(8)
Text7.Text = Data1.Recordset.Fields(9)
Text8.Text = Data1.Recordset.Fields(10)
Text6.Text = Data1.Recordset.Fields(11)
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
Combo1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)
Text3.Text = Data1.Recordset.Fields(3)
Combo2.AddItem (Data1.Recordset.Fields(4))
Combo2.Text = Data1.Recordset.Fields(4)
Text4.Text = Data1.Recordset.Fields(5)
Text5.Text = Data1.Recordset.Fields(6)
Text9.Text = Data1.Recordset.Fields(8)
Text7.Text = Data1.Recordset.Fields(9)
Text8.Text = Data1.Recordset.Fields(10)
Text6.Text = Data1.Recordset.Fields(11)
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
Combo1.Text = Data1.Recordset.Fields(1)
Text2.Text = Data1.Recordset.Fields(2)
Text3.Text = Data1.Recordset.Fields(3)
Combo2.AddItem (Data1.Recordset.Fields(4))
Combo2.Text = Data1.Recordset.Fields(4)
Text4.Text = Data1.Recordset.Fields(5)
Text5.Text = Data1.Recordset.Fields(6)
Text9.Text = Data1.Recordset.Fields(8)
Text7.Text = Data1.Recordset.Fields(9)
Text8.Text = Data1.Recordset.Fields(10)
Text6.Text = Data1.Recordset.Fields(11)
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
Combo1.ListIndex = -1
Combo2.Clear
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
'add items to combo
Combo1.Clear
If Data2.Recordset.RecordCount <> 0 Then
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Combo1.AddItem (Data2.Recordset.Fields(0))
Data2.Recordset.MoveNext
Loop
Else
MsgBox ("Purchase stock table is empty")
Hom.Show
Unload Me
End If
'combo2
Combo2.Enabled = False
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

Private Sub Text3_Change()
If Not IsNumeric(Text3.Text) Then
Text3.Text = ""
End If

Text4.Text = ""
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5.Text) Then
Text5.Text = ""
End If

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
If Not IsNumeric(Text8.Text) Then
Text8.Text = ""
End If

End Sub

Private Sub Text9_Change()
If Not IsNumeric(Text9.Text) Then
Text9.Text = ""
End If

End Sub
