VERSION 5.00
Begin VB.Form frmdeletion 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Deletion Manager"
   ClientHeight    =   3555
   ClientLeft      =   3780
   ClientTop       =   2400
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
   ScaleHeight     =   3555
   ScaleWidth      =   4860
   Begin VB.CommandButton cmdremove 
      Caption         =   "Remove"
      Height          =   255
      Left            =   2850
      TabIndex        =   11
      Top             =   3210
      Width           =   930
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   3210
      Width           =   930
   End
   Begin VB.TextBox txtaddress 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   9
      Top             =   2175
      Width           =   3030
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   8
      Top             =   1845
      Width           =   3030
   End
   Begin VB.TextBox txtsalary 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   7
      Top             =   2835
      Width           =   1710
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   6
      Top             =   2505
      Width           =   1710
   End
   Begin VB.TextBox txtempid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   5
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   4800
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Employee Deletion Manager"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   0
         Width           =   3420
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
      TabIndex        =   2
      Top             =   795
      Width           =   2430
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   0
      Top             =   1215
      Width           =   4800
      Begin VB.Label lblemp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details of Employee to be Removed"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   930
         TabIndex        =   1
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   3540
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
      TabIndex        =   17
      Top             =   2895
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   16
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   15
      Top             =   2250
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   14
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   195
      TabIndex        =   13
      Top             =   1590
      Width           =   1200
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Employee ID from the list to be deleted or select all from remove popup menu :"
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Left            =   390
      TabIndex        =   12
      Top             =   300
      Width           =   3915
   End
   Begin VB.Menu mnuremove 
      Caption         =   "Remove"
      Visible         =   0   'False
      Begin VB.Menu mnusel 
         Caption         =   "Selected"
      End
      Begin VB.Menu mnuall 
         Caption         =   "All"
      End
   End
End
Attribute VB_Name = "frmdeletion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboempid_Click()
If cboempid.ListIndex > 0 Then
    Set rs = db.OpenRecordset("select *from info where empno='" & cboempid.Text & "'")
    If rs.RecordCount > 0 Then
        lblemp.Caption = "Details of Employee " & cboempid.Text & " to be Removed"
        txtempid.Text = rs("empno")
        txtname.Text = rs("name")
        txtaddress.Text = rs("address")
        txtphone.Text = rs("phone")
        txtsalary.Text = rs("salary")
        txtempid.BackColor = vbWhite
        txtname.BackColor = vbWhite
        txtaddress.BackColor = vbWhite
        txtphone.BackColor = vbWhite
        txtsalary.BackColor = vbWhite
    End If
    Set rs = Nothing
Else
    Exit Sub
End If
End Sub

Private Sub cmdclose_Click()
confirm = MsgBox("Are you sure to exit ?", vbApplicationModal + vbYesNo + vbQuestion, "Exit Deletion")
If confirm = vbNo Then
    cboempid.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    Unload Me
    frmmenu.optadd.SetFocus
End If
End Sub

Private Sub cmdremove_Click()
If Trim(txtempid.Text) <> "" Then
    mnusel.Enabled = True
Else
    mnusel.Enabled = False
End If
If cboempid.ListCount > 1 Then
    mnuall.Enabled = True
ElseIf cboempid.ListCount = 1 Then
    mnuall.Enabled = False
End If
PopupMenu mnuremove
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2

Set db = OpenDatabase(App.Path & "\employee.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst

cboempid.AddItem "-=Select Emp ID=-"
Do Until rs.EOF
    cboempid.AddItem rs("empno")
    rs.MoveNext
Loop
cboempid.ListIndex = 0
txtempid.BackColor = &HC0C0C0
txtname.BackColor = &HC0C0C0
txtaddress.BackColor = &HC0C0C0
txtphone.BackColor = &HC0C0C0
txtsalary.BackColor = &HC0C0C0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then cmdclose_Click
End Sub

Private Sub mnuall_Click()
confirm = MsgBox("Are you sure to truncate the employee record table ?", vbApplicationModal + vbYesNo + vbQuestion, _
            "Confirm Removal")
If confirm = vbNo Then
    cboempid.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    frmdelprogress.Show vbModal
End If
End Sub

Private Sub mnusel_Click()
confirm = MsgBox("Are you sure to remove record of Employee " & txtempid.Text & " from the database ?", vbApplicationModal + vbYesNo + vbQuestion, _
            "Confirm Removal")
If confirm = vbNo Then
    txtempid.Text = ""
    txtaddress.Text = ""
    txtname.Text = ""
    txtphone.Text = ""
    txtsalary.Text = ""
    txtempid.BackColor = &HC0C0C0
    txtname.BackColor = &HC0C0C0
    txtaddress.BackColor = &HC0C0C0
    txtphone.BackColor = &HC0C0C0
    txtsalary.BackColor = &HC0C0C0
    cboempid.ListIndex = 0
    cboempid.SetFocus
ElseIf confirm = vbYes Then
    Set rs = db.OpenRecordset("select *from info where empno='" & txtempid.Text & "'")
    If rs.RecordCount > 0 Then
        rs.Delete
    End If
    txtempid.Text = ""
    txtaddress.Text = ""
    txtname.Text = ""
    txtphone.Text = ""
    txtsalary.Text = ""
    txtempid.BackColor = &HC0C0C0
    txtname.BackColor = &HC0C0C0
    txtaddress.BackColor = &HC0C0C0
    txtphone.BackColor = &HC0C0C0
    txtsalary.BackColor = &HC0C0C0
    MsgBox "Successfully removed employee record from the database.", vbApplicationModal + vbInformation, "Done"
    Set rs = db.OpenRecordset("info", dbOpenTable)
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        cboempid.Clear
        cboempid.AddItem "-=Select Emp ID=-"
        Do Until rs.EOF
            cboempid.AddItem rs("empno")
            rs.MoveNext
        Loop
        cboempid.ListIndex = 0
        cboempid.SetFocus
    ElseIf rs.RecordCount = 0 Then
        MsgBox "There is currently no employee record available for further removal." & vbCrLf & "Employee deletion manager will now close.", _
            vbApplicationModal + vbInformation, "Empty Database"
        Unload Me
        frmmenu.optadd.SetFocus
    End If
End If
End Sub
