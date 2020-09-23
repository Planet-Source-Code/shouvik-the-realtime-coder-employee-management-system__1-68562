VERSION 5.00
Begin VB.Form frmmodification 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Modification Manager"
   ClientHeight    =   3390
   ClientLeft      =   3420
   ClientTop       =   2745
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4860
   Begin VB.CommandButton cmdrequery 
      Caption         =   "Re-Query"
      Height          =   255
      Left            =   90
      TabIndex        =   19
      Top             =   3045
      Width           =   930
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   17
      Top             =   990
      Width           =   4800
      Begin VB.Label lblemp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit Details of Employee"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1410
         TabIndex        =   18
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.ComboBox cboempid 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   585
      Width           =   2430
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   9
      Top             =   30
      Width           =   4800
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Employee Modification Manager"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   555
         TabIndex        =   10
         Top             =   0
         Width           =   3750
      End
   End
   Begin VB.TextBox txtempid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   8
      Top             =   1290
      Width           =   1215
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   3
      Top             =   2280
      Width           =   1710
   End
   Begin VB.TextBox txtsalary 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   4
      Top             =   2610
      Width           =   1710
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   1
      Top             =   1620
      Width           =   3030
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1725
      TabIndex        =   2
      Top             =   1950
      Width           =   3030
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   3045
      Width           =   930
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   255
      Left            =   1860
      TabIndex        =   5
      Top             =   3045
      Width           =   930
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2850
      TabIndex        =   6
      Top             =   3045
      Width           =   930
   End
   Begin VB.Label lblreset 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reset Original Values"
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   3510
      MouseIcon       =   "frmmodification.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   2475
      Width           =   1275
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Employee ID from the list to be updated :"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   435
      TabIndex        =   16
      Top             =   315
      Width           =   4035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   15
      Top             =   1365
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   14
      Top             =   1695
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   13
      Top             =   2025
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   12
      Top             =   2355
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Salary : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   11
      Top             =   2670
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   3375
      Left            =   15
      Top             =   15
      Width           =   4845
   End
End
Attribute VB_Name = "frmmodification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboempid_Click()
If cboempid.ListIndex > 0 Then
    Set rsd = db.OpenRecordset("select *from info where empno='" & cboempid.Text & "'")
    If rsd.RecordCount > 0 Then
        lblemp.Caption = "Edit Details of Employee " & cboempid.Text
        txtempid.Text = rsd("empno")
        txtname.Text = rsd("name")
        txtaddress.Text = rsd("address")
        txtphone.Text = rsd("phone")
        txtsalary.Text = rsd("salary")
        cmdrequery.Enabled = True
        Label6.Enabled = False
        cboempid.Enabled = False
        cmdupdate.Enabled = True
        cmdclear.Enabled = True
        txtaddress.Enabled = True
        txtname.Enabled = True
        txtname.SetFocus
        txtphone.Enabled = True
        txtsalary.Enabled = True
        Label1.Enabled = True
        Label2.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        Label5.Enabled = True
        txtempid.BackColor = vbWhite
        txtname.BackColor = vbWhite
        txtaddress.BackColor = vbWhite
        txtphone.BackColor = vbWhite
        txtsalary.BackColor = vbWhite
        lblreset.Visible = True
    End If
    Set rsd = Nothing
Else
    Exit Sub
End If
End Sub

Private Sub cmdclear_Click()
txtname.Text = ""
txtaddress.Text = ""
txtphone.Text = ""
txtsalary.Text = ""
txtname.SetFocus
End Sub

Private Sub cmdclose_Click()
confirm = MsgBox("Are you sure to exit ?" & Chr(10) & "Any unsaved data will be lost.", vbApplicationModal + vbYesNo + vbQuestion, "Exit Modification")
If confirm = vbNo Then
    If txtname.Enabled = True Then
        txtname.SetFocus
    Else
        cboempid.SetFocus
    End If
    Exit Sub
ElseIf confirm = vbYes Then
    Unload Me
    frmmenu.optadd.SetFocus
End If
End Sub

Private Sub cmdrequery_Click()
Label6.Enabled = True
cboempid.Enabled = True
cboempid.ListIndex = 0
cmdrequery.Enabled = False
lblemp.Caption = "Employee Data Updation"
cmdclear_Click
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
cmdupdate.Enabled = False
cmdclear.Enabled = False
txtempid.Text = ""
txtaddress.Enabled = False
txtname.Enabled = False
txtphone.Enabled = False
txtsalary.Enabled = False
txtempid.BackColor = &HC0C0C0
txtname.BackColor = &HC0C0C0
txtaddress.BackColor = &HC0C0C0
txtphone.BackColor = &HC0C0C0
txtsalary.BackColor = &HC0C0C0
lblreset.Visible = False
cboempid.SetFocus
End Sub

Private Sub cmdupdate_Click()
If Trim(txtname.Text) <> "" And Trim(txtaddress.Text) <> "" And Trim(txtphone.Text) <> "" And Trim(txtsalary.Text) <> "" Then
    confirm = MsgBox("Are you sure to update this record ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Update")
    If confirm = vbNo Then
        txtname.SetFocus
        Exit Sub
    ElseIf confirm = vbYes Then
        Set rs = db.OpenRecordset("select *from info where empno='" & Trim(txtempid.Text) & "'")
        If rs.RecordCount > 0 Then
            rs.Edit
        End If
        rs("empno") = txtempid.Text
        rs("name") = Trim(txtname.Text)
        rs("address") = Trim(txtaddress.Text)
        rs("phone") = Trim(txtphone.Text)
        rs("salary") = Trim(txtsalary.Text)
        rs.Update
        MsgBox "Successfully updated database.", vbApplicationModal + vbInformation, "Done"
        cmdrequery_Click
        Set rs = Nothing
    End If
Else
    MsgBox "Sorry,modification manager cannot" & Chr(10) & "update the database with blank values.", vbApplicationModal + vbExclamation, _
        "Incomplete Information"
    txtname.SetFocus
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

cboempid.AddItem "-=Select Emp ID=-"
Do Until rs.EOF
    cboempid.AddItem rs("empno")
    rs.MoveNext
Loop
cboempid.ListIndex = 0
lblemp.Caption = "Employee Data Updation"
cmdrequery.Enabled = False
cmdupdate.Enabled = False
cmdclear.Enabled = False
txtaddress.Enabled = False
txtname.Enabled = False
txtphone.Enabled = False
txtsalary.Enabled = False
Label1.Enabled = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
txtempid.BackColor = &HC0C0C0
txtname.BackColor = &HC0C0C0
txtaddress.BackColor = &HC0C0C0
txtphone.BackColor = &HC0C0C0
txtsalary.BackColor = &HC0C0C0
lblreset.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblreset.FontUnderline = False
End Sub

Private Sub lblreset_Click()
Set rsd = db.OpenRecordset("select *from info where empno='" & cboempid.Text & "'")
If rsd.RecordCount > 0 Then
    lblemp.Caption = "Edit Details of Employee " & cboempid.Text
    txtempid.Text = rsd("empno")
    txtname.Text = rsd("name")
    txtaddress.Text = rsd("address")
    txtphone.Text = rsd("phone")
    txtsalary.Text = rsd("salary")
End If
Set rsd = Nothing
End Sub

Private Sub lblreset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblreset.FontUnderline = True
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
