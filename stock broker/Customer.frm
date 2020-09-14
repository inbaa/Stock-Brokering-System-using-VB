VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Customer 
   Caption         =   "Customer"
   ClientHeight    =   9660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9525
   LinkTopic       =   "Form2"
   ScaleHeight     =   9660
   ScaleWidth      =   9525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   4800
      TabIndex        =   26
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2880
      TabIndex        =   27
      Top             =   8760
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      Caption         =   "BACK"
      Height          =   615
      Left            =   7440
      MaskColor       =   &H0080FFFF&
      TabIndex        =   22
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SAVE"
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
      Left            =   7320
      TabIndex        =   21
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7320
      TabIndex        =   20
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   7320
      TabIndex        =   19
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text13 
      DataField       =   "lanline"
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   7200
      Width           =   3015
   End
   Begin VB.TextBox Text12 
      DataField       =   "phone_no2"
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Top             =   6720
      Width           =   3015
   End
   Begin VB.TextBox Text11 
      DataField       =   "phone_no1"
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Top             =   6240
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "supp_name"
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "area"
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "street"
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "house_no"
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   240
      TabIndex        =   29
      Top             =   600
      Width           =   6615
      Begin VB.TextBox Text1 
         DataField       =   "supp_id"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         DataField       =   "city"
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         DataField       =   "state"
         Height          =   285
         Left            =   2640
         TabIndex        =   7
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         DataField       =   "country"
         Height          =   285
         Left            =   2640
         TabIndex        =   8
         Top             =   4080
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         DataField       =   "pincode"
         Height          =   285
         Left            =   2640
         TabIndex        =   9
         Top             =   4560
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         DataField       =   "zipcode"
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   5040
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Customer.frx":0000
         Left            =   5040
         List            =   "Customer.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   7080
         Width           =   1335
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   7080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Cutomer ID"
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
         TabIndex        =   45
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Address :"
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
         TabIndex        =   44
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "House No"
         Height          =   255
         Left            =   720
         TabIndex        =   43
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Street"
         Height          =   255
         Left            =   720
         TabIndex        =   42
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Area"
         Height          =   255
         Left            =   720
         TabIndex        =   41
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "City"
         Height          =   255
         Left            =   720
         TabIndex        =   40
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "State"
         Height          =   375
         Left            =   720
         TabIndex        =   39
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Country"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Pincode"
         Height          =   255
         Left            =   720
         TabIndex        =   37
         Top             =   4560
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Zip Code"
         Height          =   255
         Left            =   720
         TabIndex        =   36
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Mobile Number"
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
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label Label14 
         Caption         =   "Alternate Mobile Number"
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
         TabIndex        =   34
         Top             =   6240
         Width           =   2175
      End
      Begin VB.Label Label15 
         Caption         =   "Lanline Number"
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
         TabIndex        =   33
         Top             =   6720
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Email "
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
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "@"
         Height          =   255
         Left            =   4800
         TabIndex        =   31
         Top             =   7080
         Width           =   255
      End
      Begin VB.Label Label20 
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
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   28
      Top             =   8400
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "ADD NEW"
         Height          =   495
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4095
      Left            =   7080
      TabIndex        =   24
      Top             =   600
      Width           =   2175
      Begin VB.CommandButton Command5 
         Caption         =   "FIRST"
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   840
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
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3375
      Left            =   7080
      TabIndex        =   23
      Top             =   6120
      Width           =   2175
      Begin VB.CommandButton Command9 
         Caption         =   "CLEAR"
         Height          =   615
         Left            =   360
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
      Height          =   855
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "customer"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label21 
      Caption         =   "Customer information"
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
      TabIndex        =   49
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Supplier Name"
      Height          =   255
      Left            =   720
      TabIndex        =   48
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Left            =   600
      TabIndex        =   47
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label19 
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
      Left            =   600
      TabIndex        =   46
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


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
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
MaskEdBox1.Text = ""
Combo1.ListIndex = -1
'AUTO GENERATE
Data1.Refresh
v = Data1.Recordset.RecordCount
If v = 0 Then
  Text1.Text = "CUST_1"
  Else
  Data1.Recordset.MoveLast
  p = Data1.Recordset.Fields(0)
  num = Mid(p, 6)
  res = num + 1
  Text1.Text = "CUST_" & res
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
n1 = Len(Text11.Text)
n2 = Len(Text12.Text)

If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or MaskEdBox1.Text = "" Or Combo1.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
ElseIf (n1 < 10) And (n2 < 10) Then
MsgBox ("Mobiler number is not valid")
Else
Data1.Refresh
Data1.Recordset.AddNew
Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Val(Text3.Text)
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = Text5.Text
Data1.Recordset.Fields(5) = Text6.Text
Data1.Recordset.Fields(6) = Text7.Text
Data1.Recordset.Fields(7) = Text8.Text
Data1.Recordset.Fields(8) = Val(Text9.Text)
Data1.Recordset.Fields(9) = Val(Text10.Text)
Data1.Recordset.Fields(10) = Val(Text11.Text)
Data1.Recordset.Fields(11) = Val(Text12.Text)
Data1.Recordset.Fields(12) = Val(Text13.Text)
Data1.Recordset.Fields(13) = MaskEdBox1.Text + "@" + Combo1.Text
Data1.Recordset.Update
MsgBox ("Customer information saved successfully")
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
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
 If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
 Text1.Text = Data1.Recordset.Fields(0)
 Text2.Text = Data1.Recordset.Fields(1)
 Text3.Text = Data1.Recordset.Fields(2)
 Text4.Text = Data1.Recordset.Fields(3)
 Text5.Text = Data1.Recordset.Fields(4)
 Text6.Text = Data1.Recordset.Fields(5)
 Text7.Text = Data1.Recordset.Fields(6)
 Text8.Text = Data1.Recordset.Fields(7)
 Text9.Text = Data1.Recordset.Fields(8)
 Text10.Text = Data1.Recordset.Fields(9)
 Text11.Text = Data1.Recordset.Fields(10)
 Text12.Text = Data1.Recordset.Fields(11)
 Text13.Text = Data1.Recordset.Fields(12)
 s = Data1.Recordset.Fields(13)
 pos = InStr(s, "@")
 l = Len(s)
 MaskEdBox1.Text = Left(s, pos - 1)
 Combo1.Text = Right(s, l - pos)
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
Command3.Enabled = False
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
flag = 0
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
 If (LCase(a) = LCase(Data1.Recordset.Fields(0))) Then
 Text1.Text = Data1.Recordset.Fields(0)
 Text2.Text = Data1.Recordset.Fields(1)
 Text3.Text = Data1.Recordset.Fields(2)
 Text4.Text = Data1.Recordset.Fields(3)
 Text5.Text = Data1.Recordset.Fields(4)
 Text6.Text = Data1.Recordset.Fields(5)
 Text7.Text = Data1.Recordset.Fields(6)
 Text8.Text = Data1.Recordset.Fields(7)
 Text9.Text = Data1.Recordset.Fields(8)
 Text10.Text = Data1.Recordset.Fields(9)
 Text11.Text = Data1.Recordset.Fields(10)
 Text12.Text = Data1.Recordset.Fields(11)
 Text13.Text = Data1.Recordset.Fields(12)
 s = Data1.Recordset.Fields(13)
 pos = InStr(s, "@")
 l = Len(s)
 MaskEdBox1.Text = Left(s, pos - 1)
 Combo1.Text = Right(s, l - pos)
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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
MaskEdBox1.Text = ""
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
n1 = Len(Text11.Text)
n2 = Len(Text12.Text)

If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or MaskEdBox1.Text = "" Or Combo1.Text = "" Then
a = MsgBox("Please fill all the contents!", vbExclamation)
ElseIf (n1 < 10) And (n2 < 10) Then
MsgBox ("Mobiler number is not valid")
Else

Data1.Recordset.Fields(0) = Text1.Text
Data1.Recordset.Fields(1) = Text2.Text
Data1.Recordset.Fields(2) = Val(Text3.Text)
Data1.Recordset.Fields(3) = Text4.Text
Data1.Recordset.Fields(4) = Text5.Text
Data1.Recordset.Fields(5) = Text6.Text
Data1.Recordset.Fields(6) = Text7.Text
Data1.Recordset.Fields(7) = Text8.Text
Data1.Recordset.Fields(8) = Val(Text9.Text)
Data1.Recordset.Fields(9) = Val(Text10.Text)
Data1.Recordset.Fields(10) = Val(Text11.Text)
Data1.Recordset.Fields(11) = Val(Text12.Text)
Data1.Recordset.Fields(12) = Val(Text13.Text)
Data1.Recordset.Fields(13) = MaskEdBox1.Text + "@" + Combo1.Text
Data1.Recordset.Update
MsgBox ("Updated successfully")
Frame3.Enabled = True
Frame4.Enabled = True
Command1.Enabled = True
Command3.Enabled = True
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
 Text2.Text = Data1.Recordset.Fields(1)
 Text3.Text = Data1.Recordset.Fields(2)
 Text4.Text = Data1.Recordset.Fields(3)
 Text5.Text = Data1.Recordset.Fields(4)
 Text6.Text = Data1.Recordset.Fields(5)
 Text7.Text = Data1.Recordset.Fields(6)
 Text8.Text = Data1.Recordset.Fields(7)
 Text9.Text = Data1.Recordset.Fields(8)
 Text10.Text = Data1.Recordset.Fields(9)
 Text11.Text = Data1.Recordset.Fields(10)
 Text12.Text = Data1.Recordset.Fields(11)
 Text13.Text = Data1.Recordset.Fields(12)
 s = Data1.Recordset.Fields(13)
 pos = InStr(s, "@")
 l = Len(s)
 MaskEdBox1.Text = Left(s, pos - 1)
 Combo1.Text = Right(s, l - pos)
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
 Text4.Text = Data1.Recordset.Fields(3)
 Text5.Text = Data1.Recordset.Fields(4)
 Text6.Text = Data1.Recordset.Fields(5)
 Text7.Text = Data1.Recordset.Fields(6)
 Text8.Text = Data1.Recordset.Fields(7)
 Text9.Text = Data1.Recordset.Fields(8)
 Text10.Text = Data1.Recordset.Fields(9)
 Text11.Text = Data1.Recordset.Fields(10)
 Text12.Text = Data1.Recordset.Fields(11)
 Text13.Text = Data1.Recordset.Fields(12)
 s = Data1.Recordset.Fields(13)
 pos = InStr(s, "@")
 l = Len(s)
 MaskEdBox1.Text = Left(s, pos - 1)
 Combo1.Text = Right(s, l - pos)
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
 Text4.Text = Data1.Recordset.Fields(3)
 Text5.Text = Data1.Recordset.Fields(4)
 Text6.Text = Data1.Recordset.Fields(5)
 Text7.Text = Data1.Recordset.Fields(6)
 Text8.Text = Data1.Recordset.Fields(7)
 Text9.Text = Data1.Recordset.Fields(8)
 Text10.Text = Data1.Recordset.Fields(9)
 Text11.Text = Data1.Recordset.Fields(10)
 Text12.Text = Data1.Recordset.Fields(11)
 Text13.Text = Data1.Recordset.Fields(12)
 s = Data1.Recordset.Fields(13)
 pos = InStr(s, "@")
 l = Len(s)
 MaskEdBox1.Text = Left(s, pos - 1)
 Combo1.Text = Right(s, l - pos)
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
 Text2.Text = Data1.Recordset.Fields(1)
 Text3.Text = Data1.Recordset.Fields(2)
 Text4.Text = Data1.Recordset.Fields(3)
 Text5.Text = Data1.Recordset.Fields(4)
 Text6.Text = Data1.Recordset.Fields(5)
 Text7.Text = Data1.Recordset.Fields(6)
 Text8.Text = Data1.Recordset.Fields(7)
 Text9.Text = Data1.Recordset.Fields(8)
 Text10.Text = Data1.Recordset.Fields(9)
 Text11.Text = Data1.Recordset.Fields(10)
 Text12.Text = Data1.Recordset.Fields(11)
 Text13.Text = Data1.Recordset.Fields(12)
 s = Data1.Recordset.Fields(13)
 pos = InStr(s, "@")
 l = Len(s)
 MaskEdBox1.Text = Left(s, pos - 1)
 Combo1.Text = Right(s, l - pos)
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
Text11.Text = ""
Text12.Text = ""
Text13.Text = ""
MaskEdBox1.Text = ""
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
'Save
Command4.Enabled = False
Command10.Enabled = False
End Sub

Private Sub Text10_Change()
If Not IsNumeric(Text10.Text) Then
Text10.Text = ""
End If
End Sub

Private Sub Text11_Change()
If Not IsNumeric(Text11.Text) Then
Text11.Text = ""
End If
End Sub

Private Sub Text12_Change()
If Not IsNumeric(Text12.Text) Then
Text12.Text = ""
End If
End Sub

Private Sub Text13_Change()
If Not IsNumeric(Text13.Text) Then
Text13.Text = ""
End If
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
If IsNumeric(Text4.Text) Then
Text4.Text = ""
End If
End Sub

Private Sub Text5_Change()
If IsNumeric(Text5.Text) Then
Text5.Text = ""
End If
End Sub

Private Sub Text6_Change()
If IsNumeric(Text6.Text) Then
Text6.Text = ""
End If
End Sub

Private Sub Text7_Change()
If IsNumeric(Text7.Text) Then
Text7.Text = ""
End If
End Sub

Private Sub Text8_Change()
If IsNumeric(Text8.Text) Then
Text8.Text = ""
End If
End Sub

Private Sub Text9_Change()
If Not IsNumeric(Text9.Text) Then
Text9.Text = ""
End If
End Sub
