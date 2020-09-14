VERSION 5.00
Begin VB.Form Supplier_Stock 
   Caption         =   "Supplier Stock"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   24
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   4680
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
      Height          =   420
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "supplier"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7080
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7200
      MaskColor       =   &H0080FFFF&
      TabIndex        =   9
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   7200
      MaskColor       =   &H0080FFFF&
      TabIndex        =   4
      Top             =   4800
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Supplier_Stock.frx":0000
      Left            =   2880
      List            =   "Supplier_Stock.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\III year Project\PROJECT.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "supplier_stock"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   6840
      TabIndex        =   22
      Top             =   4200
      Width           =   2175
      Begin VB.CommandButton Command9 
         Caption         =   "CLEAR"
         Height          =   615
         Left            =   360
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   6840
      TabIndex        =   19
      Top             =   480
      Width           =   2175
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000017&
      Height          =   1095
      Left            =   120
      TabIndex        =   18
      Top             =   5520
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "ADD NEW"
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   6495
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         Top             =   3360
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Supplier_Stock.frx":0004
         Left            =   2760
         List            =   "Supplier_Stock.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   2760
         Width           =   1815
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
         TabIndex        =   25
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label20 
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
         Left            =   360
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label5 
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
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "@"
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   7080
         Width           =   255
      End
   End
   Begin VB.Label Label21 
      Caption         =   "Supplier Stock"
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
      Left            =   2760
      TabIndex        =   23
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Supplier_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo2_Click()
'sup name
t = Combo1.Text
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
If (t = Data2.Recordset.Fields(0)) Then
Text1.Text = Data2.Recordset.Fields(1)
Exit Do
Else
Data2.Recordset.MoveNext
End If
Loop
'rate empty
Text3.Text = ""
End Sub

Private Sub Command1_Click()
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
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
'to check duplicate in supp_id and Pitem_no
f = 0
Data1.Refresh
If Combo1.Text = "" Or Combo2.Text = "" Or Text3.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
f = 1
ElseIf Val(Text3.Text) <= 0 Then
a = MsgBox("Rate cannot be less than 0!", vbExclamation)
Text3.Text = ""
f = 1
ElseIf f = 0 Then
 supp = Combo1.Text
 pitem = Combo2.Text
 Do While Not Data1.Recordset.EOF
 If supp = Data1.Recordset.Fields(0) And pitem = Data1.Recordset.Fields(2) Then
 b = MsgBox("Given Supplier already supplied this item!", vbExclamation)
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
Data1.Recordset.Fields(0) = Combo1.Text
Data1.Recordset.Fields(1) = Text1.Text
Data1.Recordset.Fields(2) = Combo2.Text
Data1.Recordset.Fields(3) = Text2.Text
Data1.Recordset.Fields(4) = Val(Text3.Text)
Data1.Recordset.Update
MsgBox ("Data saved successfully")
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

Private Sub Command2_Click()
a = InputBox("Enter Supplier ID")
b = InputBox("Enter Purchase Item Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
 If (LCase(a) = LCase(Data1.Recordset.Fields(0)) And LCase(b) = LCase(Data1.Recordset.Fields(2))) Then
  Combo1.Text = Data1.Recordset.Fields(0)
  Text1.Text = Data1.Recordset.Fields(1)
  Combo2.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
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
a = InputBox("Enter Supplier ID")
b = InputBox("Enter Purchase Item Number")

flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0)) And LCase(b) = LCase(Data1.Recordset.Fields(2))) Then
  Combo1.Text = Data1.Recordset.Fields(0)
  Text1.Text = Data1.Recordset.Fields(1)
  Combo2.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
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
  'Clear code
 Combo1.ListIndex = -1
Combo2.ListIndex = -1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""

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
'to check duplicate in supp_id and Pitem_no
f = 0
Data1.Refresh
If Combo1.Text = "" Or Combo2.Text = "" Or Text3.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
f = 1
ElseIf Val(Text3.Text) <= 0 Then
a = MsgBox("Rate cannot be less than 0!", vbExclamation)
Text3.Text = ""
f = 1
ElseIf f = 0 Then
 supp = Combo1.Text
 pitem = Combo2.Text
 Do While Not Data1.Recordset.EOF
 If supp = Data1.Recordset.Fields(0) And pitem = Data1.Recordset.Fields(2) Then
 b = MsgBox("Given Supplier already supplied this item!", vbExclamation)
 f = 1
 Exit Do
 Else
 Data1.Recordset.MoveNext
 End If
 Loop
End If
If f = 0 Then
Data1.Recordset.Fields(0) = Combo1.Text
Data1.Recordset.Fields(1) = Text1.Text
Data1.Recordset.Fields(2) = Combo2.Text
Data1.Recordset.Fields(3) = Text2.Text
Data1.Recordset.Fields(4) = Text3.Text
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
 Combo1.Text = Data1.Recordset.Fields(0)
  Text1.Text = Data1.Recordset.Fields(1)
  Combo2.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
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
 Combo1.Text = Data1.Recordset.Fields(0)
  Text1.Text = Data1.Recordset.Fields(1)
  Combo2.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
 Command7.Enabled = True
 End If
End Sub

Private Sub Command7_Click()
Data1.Recordset.MoveNext
If (Data1.Recordset.EOF = True) Then
 MsgBox ("This is last record")
 Command7.Enabled = False
 Else
Combo1.Text = Data1.Recordset.Fields(0)
  Text1.Text = Data1.Recordset.Fields(1)
  Combo2.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
 Command6.Enabled = True
 End If

End Sub

Private Sub Command8_Click()
 Data1.Recordset.MoveLast
 If Data1.Recordset.BOF = True Then
MsgBox ("No record to display")
Command8.Enabled = False
Else
Combo1.Text = Data1.Recordset.Fields(0)
  Text1.Text = Data1.Recordset.Fields(1)
  Combo2.Text = Data1.Recordset.Fields(2)
  Text2.Text = Data1.Recordset.Fields(3)
  Text3.Text = Data1.Recordset.Fields(4)
 Command7.Enabled = False
 Command6.Enabled = True
 End If

End Sub

Private Sub Command9_Click()
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
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
If Data2.Recordset.RecordCount <> 0 Then
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
Combo1.AddItem (Data2.Recordset.Fields(0))
Data2.Recordset.MoveNext
Loop
Else
MsgBox ("Supplier info table is empty")
Hom.Show
Unload Me
End If

If Data3.Recordset.RecordCount <> 0 Then
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
Combo2.AddItem (Data3.Recordset.Fields(0))
Data3.Recordset.MoveNext
Loop
Else
MsgBox ("Purchase item table is empty")
Hom.Show
Unload Me
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

Private Sub Text3_Change()
'validation
If Not IsNumeric(Text3.Text) Then
Text3.Text = ""
End If

End Sub

Private Sub Text3_GotFocus()
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

End Sub
