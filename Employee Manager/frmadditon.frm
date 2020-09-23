VERSION 5.00
Begin VB.Form frmadditon 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Addition Manager"
   ClientHeight    =   2490
   ClientLeft      =   3585
   ClientTop       =   3030
   ClientWidth     =   4860
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmadditon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4860
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2850
      TabIndex        =   6
      Top             =   2145
      Width           =   930
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   255
      Left            =   1860
      TabIndex        =   5
      Top             =   2145
      Width           =   930
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   2145
      Width           =   930
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   2
      Top             =   1020
      Width           =   3030
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   1
      Top             =   690
      Width           =   3030
   End
   Begin VB.TextBox txtsalary 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   4
      Top             =   1680
      Width           =   1710
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   3
      Top             =   1350
      Width           =   1710
   End
   Begin VB.TextBox txtempid 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   13
      Top             =   30
      Width           =   4800
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Employee Addition Manager"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   14
         Top             =   0
         Width           =   3420
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   2475
      Left            =   15
      Top             =   15
      Width           =   4845
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Salary : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   12
      Top             =   1740
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   11
      Top             =   1425
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   10
      Top             =   1095
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   9
      Top             =   765
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   8
      Top             =   435
      Width           =   1200
   End
End
Attribute VB_Name = "frmadditon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdclear_Click()
txtempid.Text = ""
txtname.Text = ""
txtaddress.Text = ""
txtphone.Text = ""
txtsalary.Text = ""
txtempid.SetFocus
End Sub

Private Sub cmdclose_Click()
confirm = MsgBox("Are you sure to exit ?" & Chr(10) & "Any unsaved data will be lost.", vbApplicationModal + vbYesNo + vbQuestion, "Exit Addition")
If confirm = vbNo Then
    txtempid.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    Unload Me
    frmmenu.optadd.SetFocus
End If
End Sub

Private Sub cmdsave_Click()
If Trim(txtempid.Text) <> "" And Trim(txtname.Text) <> "" And Trim(txtaddress.Text) <> "" And Trim(txtphone.Text) <> "" _
    And Trim(txtsalary.Text) <> "" Then
    Set rsd = db.OpenRecordset("select * from info where empno='" & Trim(txtempid.Text) & "'")
    If rsd.RecordCount > 0 Then
        MsgBox "Employee " & Chr(34) & Trim(txtempid.Text) & Chr(34) & " already exists.", vbApplicationModal + vbInformation, "Duplication Found"
        txtempid.SelStart = 0
        txtempid.SelLength = Len(Trim(txtempid.Text))
        txtempid.SetFocus
        Exit Sub
    ElseIf rsd.RecordCount = 0 Then
        rs.AddNew
        If rs.EditMode = dbEditAdd Then
            rs("empno") = Trim(txtempid.Text)
            rs("name") = Trim(txtname.Text)
            rs("address") = Trim(txtaddress.Text)
            rs("phone") = Trim(txtphone.Text)
            rs("salary") = Trim(txtsalary.Text)
            rs.Update
            MsgBox "Successfully added new employee data.", vbApplicationModal + vbInformation, "Done"
            cmdclear_Click
        End If
    End If
    Set rsd = Nothing
Else
    MsgBox "Sorry,blank values cannot be saved.", vbApplicationModal + vbExclamation, "Incompleted Information"
    txtempid.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdclose_Click
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2

Set db = OpenDatabase(App.Path & "\employee.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
Set rsd = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst
If rsd.RecordCount > 0 Then rsd.MoveFirst
End Sub

Private Sub txtname_LostFocus()
Dim i, c As Integer, n As String

If Trim(txtname.Text) <> "" Then
    For i = 1 To Len(Trim(txtname.Text)) Step 1
        n = Mid(Trim(txtname.Text), i, 1)
        n = LCase(n)
        If n >= "a" And n <= "z" Or n = " " Then
            c = 0
        Else
            c = 1
        End If
    Next i
    If c = 0 Then
        Exit Sub
    ElseIf c = 1 Then
        MsgBox "Name contains some junk characters." & vbCrLf & "Please rectify.", vbApplicationModal + vbExclamation, "Error"
        txtname.SelStart = 0
        txtname.SelLength = Len(Trim(txtname.Text))
        txtname.SetFocus
    End If
Else
    Exit Sub
End If
End Sub

Private Sub txtphone_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtsalary_KeyPress(KeyAscii As Integer)
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

Private Sub txtphone_LostFocus()
If Trim(txtphone.Text) <> "" Then
    If InStr(1, txtphone.Text, ".") Then
        MsgBox "Phone number contains a junk character." & vbCrLf & "Please rectify.", vbApplicationModal + vbExclamation, "Error"
        txtphone.SelStart = 0
        txtphone.SelLength = Len(Trim(txtphone.Text))
        txtphone.SetFocus
    Else
        If Len(Trim(txtphone.Text)) = 8 Or Len(Trim(txtphone.Text)) = 10 Then
            Exit Sub
        Else
            MsgBox "Phone no should be consist of 8(Cell) or 10(Land) digits.", vbApplicationModal + vbInformation, "Invalid Phone No Format"
            txtphone.SelStart = 0
            txtphone.SelLength = Len(Trim(txtphone.Text))
            txtphone.SetFocus
        End If
    End If
Else
    Exit Sub
End If
End Sub

Private Sub txtsalary_LostFocus()
If Trim(txtsalary.Text) <> "" Then
    If IsNumeric(Trim(txtsalary.Text)) Then
        Exit Sub
    Else
        MsgBox "Invalid salary.", vbApplicationModal + vbExclamation, "Error"
        txtsalary.SelStart = 0
        txtsalary.SelLength = Len(Trim(txtsalary.Text))
        txtsalary.SetFocus
    End If
Else
    Exit Sub
End If
End Sub
