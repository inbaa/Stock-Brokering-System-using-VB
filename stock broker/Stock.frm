VERSION 5.00
Begin VB.Form Stock 
   Caption         =   "Stock"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form3"
   ScaleHeight     =   7995
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      DataField       =   "street"
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "house_no"
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "supp_name"
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7320
      MaskColor       =   &H0080FFFF&
      TabIndex        =   12
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SAVE"
      Height          =   615
      Left            =   7320
      MaskColor       =   &H0080FFFF&
      TabIndex        =   6
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "LAST"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7200
      TabIndex        =   10
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   6735
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Stock.frx":0000
         Left            =   2640
         List            =   "Stock.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label Label18 
         Caption         =   "@"
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label Label1 
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
         TabIndex        =   24
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Minimum Level"
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
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Unit of Measurement"
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
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label20 
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
         TabIndex        =   20
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000004&
      ForeColor       =   &H80000017&
      Height          =   1095
      Left            =   240
      TabIndex        =   18
      Top             =   5640
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "ADD NEW"
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3615
      Left            =   6960
      TabIndex        =   14
      Top             =   600
      Width           =   2175
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   8
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
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   6960
      TabIndex        =   13
      Top             =   4440
      Width           =   2175
      Begin VB.CommandButton Command9 
         Caption         =   "CLEAR"
         Height          =   615
         Left            =   360
         MaskColor       =   &H00C0FFFF&
         TabIndex        =   7
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
      Height          =   735
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "purchase_stock"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "Purchase Stock"
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
      Left            =   2880
      TabIndex        =   26
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label19_Click()

End Sub



Private Sub Command1_Click()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
 Combo1.ListIndex = -1
'AUTO GENERATE
 Data1.Refresh
 v = Data1.Recordset.RecordCount
 If v = 0 Then
  Text1.Text = "Pitem_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 7)
  res = num + 1
  Text1.Text = "Pitem_" & res
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
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
ElseIf Val(Text4.Text) > Val(Text3.Text) Then
MsgBox ("Minimum level should be less than quantity")
ElseIf Val(Text3.Text) <= 0 Then
Text4.Text = ""
MsgBox ("Quantity cannot be less than 0")
Text3.Text = ""
ElseIf Val(Text4.Text) <= 0 Then
MsgBox ("Minimum level cannot be less than 0")
Text4.Text = ""
Else
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Val(Text3.Text)
Data1.Recordset.Fields(3) = Combo1.Text
Data1.Recordset.Fields(4) = Val(Text4.Text)
Data1.Recordset.Fields(5) = 0
Data1.Recordset.Update
MsgBox ("Purchase stock details saved successfully")
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
a = InputBox("Enter Purchase Item Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
Text1.Text = Data1.Recordset.Fields(0)
Text2.Text = Data1.Recordset.Fields(1)
Text3.Text = Data1.Recordset.Fields(2)
Text3.Enabled = False
Combo1.Text = Data1.Recordset.Fields(3)
Text4.Text = Data1.Recordset.Fields(4)
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
a = InputBox("Enter Purchase Item Number")
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
Text1.Text = Data1.Recordset.Fields(0)
Text2.Text = Data1.Recordset.Fields(1)
Text3.Text = Data1.Recordset.Fields(2)
Combo1.Text = Data1.Recordset.Fields(3)
Text4.Text = Data1.Recordset.Fields(4)
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
  Combo1.ListIndex = -1
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
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
ElseIf Val(Text4.Text) > Val(Text3.Text) Then
MsgBox ("Minimum level should be less than quantity")
Text4.Text = ""
ElseIf Val(Text3.Text) <= 0 Then
MsgBox ("Quantity cannot be less than 0")
Text3.Text = ""
ElseIf Val(Text4.Text) <= 0 Then
MsgBox ("Minimum level cannot be less than 0")
Text4.Text = ""
Else

Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Val(Text3.Text)
Text3.Enabled = True
Data1.Recordset.Fields(3) = Combo1.Text
Data1.Recordset.Fields(4) = Val(Text4.Text)
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
Else
 Text1.Text = Data1.Recordset.Fields(0)
Text2.Text = Data1.Recordset.Fields(1)
Text3.Text = Data1.Recordset.Fields(2)
Combo1.Text = Data1.Recordset.Fields(3)
Text4.Text = Data1.Recordset.Fields(4)
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
Text2.Text = Data1.Recordset.Fields(1)
Text3.Text = Data1.Recordset.Fields(2)
Combo1.Text = Data1.Recordset.Fields(3)
Text4.Text = Data1.Recordset.Fields(4)
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
Text2.Text = Data1.Recordset.Fields(1)
Text3.Text = Data1.Recordset.Fields(2)
Combo1.Text = Data1.Recordset.Fields(3)
Text4.Text = Data1.Recordset.Fields(4)
 Command6.Enabled = True
 End If
End Sub

Private Sub Command8_Click()
Data1.Recordset.MoveLast
 If Data1.Recordset.BOF = True Then
MsgBox ("No record to display")
Else
Text1.Text = Data1.Recordset.Fields(0)
Text2.Text = Data1.Recordset.Fields(1)
Text3.Text = Data1.Recordset.Fields(2)
Combo1.Text = Data1.Recordset.Fields(3)
Text4.Text = Data1.Recordset.Fields(4)
 Command7.Enabled = False
 Command6.Enabled = True
 End If
End Sub


Private Sub Command9_Click()
Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 Text4.Text = ""
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
'save
Command10.Enabled = False
Command1.Enabled = True
 Command11.Enabled = True

End Sub

Private Sub Form_Activate()
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
Command10.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Text2_Change()
If IsNumeric(Text2.Text) Then
Text2.Text = ""
End If
End Sub

Private Sub Text3_Change()
If Not IsNumeric(Text3.Text) Then
Text3.Text = ""
End If
End Sub

Private Sub Text4_Change()
If Not IsNumeric(Text4.Text) Then
Text4.Text = ""
End If
End Sub

Private Sub Text5_Change()
If Not IsNumeric(Text5.Text) Then
Text5.Text = ""
End If
End Sub
