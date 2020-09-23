VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsearching 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Searching Manager"
   ClientHeight    =   3615
   ClientLeft      =   3330
   ClientTop       =   2490
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
   ScaleHeight     =   3615
   ScaleWidth      =   4860
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3660
      Top             =   1545
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   3465
      Top             =   2145
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   18
      Top             =   2880
      Width           =   4800
      Begin VB.Label lblwait 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3765
         TabIndex        =   20
         Top             =   0
         Width           =   990
      End
      Begin VB.Label lblsearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Searching"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   45
         TabIndex        =   19
         Top             =   15
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find"
      Height          =   255
      Left            =   3750
      TabIndex        =   1
      Top             =   345
      Width           =   780
   End
   Begin VB.TextBox txtsearch 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2025
      TabIndex        =   0
      Top             =   330
      Width           =   1650
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   2550
      Width           =   930
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   4800
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Employee Searching Manager"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   645
         TabIndex        =   6
         Top             =   0
         Width           =   3570
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   3
      Top             =   735
      Width           =   4800
      Begin VB.Label lblemp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details of Employee"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1590
         TabIndex        =   4
         Top             =   0
         Width           =   1680
      End
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   135
      Left            =   510
      TabIndex        =   21
      Top             =   3195
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lbldone 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Completed"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1845
      TabIndex        =   22
      Top             =   3315
      Width           =   1185
   End
   Begin VB.Label lblphone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   1905
      Width           =   660
   End
   Begin VB.Label lblsalary 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1680
      TabIndex        =   16
      Top             =   2175
      Width           =   660
   End
   Begin VB.Label lbladdress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label9"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1680
      TabIndex        =   15
      Top             =   1635
      Width           =   555
   End
   Begin VB.Label lblname 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label8"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1680
      TabIndex        =   14
      Top             =   1380
      Width           =   555
   End
   Begin VB.Label lblempid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID : "
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1680
      TabIndex        =   13
      Top             =   1110
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   3600
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
      Left            =   255
      TabIndex        =   12
      Top             =   2175
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   255
      TabIndex        =   11
      Top             =   1905
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   255
      TabIndex        =   10
      Top             =   1635
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   255
      TabIndex        =   9
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   1110
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Employee ID : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   255
      TabIndex        =   7
      Top             =   375
      Width           =   1695
   End
End
Attribute VB_Name = "frmsearching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdclose_Click()
confirm = MsgBox("Are you sure to exit ?", vbApplicationModal + vbYesNo + vbQuestion, "Exit Searching")
If confirm = vbNo Then
    txtsearch.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    Unload Me
    frmmenu.optadd.SetFocus
End If
End Sub

Private Sub cmdfind_Click()
If Trim(txtsearch.Text) <> "" Then
    lblsearch.Caption = "Searching for Employee " & txtsearch.Text
    lblwait.Caption = "Please Wait"
    Set rs = db.OpenRecordset("select *from info where empno='" & txtsearch.Text & "'")
    txtsearch.Enabled = False
    cmdfind.Enabled = False
    i = 0
    p1.Max = 100
    p1.Min = 0
    Timer1.Enabled = True
    Timer2.Enabled = True
Else
    Beep
    Exit Sub
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If txtsearch.Enabled = True Then
        cmdclose_Click
    ElseIf txtsearch.Enabled = False Then
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2
Timer1.Enabled = False
Timer2.Enabled = False

Set db = OpenDatabase(App.Path & "\employee.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst
lbldone.Caption = "0 % Completed"
lblsearch.Caption = "Search Status"
lblwait.Caption = "Stopped"
lbladdress.Caption = "-=No Information=-"
lblempid.Caption = "-=No Information=-"
lblname.Caption = "-=No Information=-"
lblphone.Caption = "-=No Information=-"
lblsalary.Caption = "-=No Information=-"
End Sub

Private Sub Timer2_Timer()
lblwait.Visible = Not lblwait.Visible
End Sub

Private Sub Timer1_Timer()
i = i + 1
p1.Value = i
lbldone.Caption = p1.Value & "% Completed"
If i = 100 Then
    If rs.RecordCount > 0 Then
        lblemp.Caption = "Details of Employee " & txtsearch.Text
        lbladdress.Caption = rs("address")
        lblempid.Caption = rs("empno")
        lblname.Caption = rs("name")
        lblphone.Caption = rs("phone")
        lblsalary.Caption = rs("salary")
    Else
        lblemp.Caption = "Details of Employee " & txtsearch.Text
        lbladdress.Caption = "-=No Information=-"
        lblempid.Caption = "-=No Information=-"
        lblname.Caption = "-=No Information=-"
        lblphone.Caption = "-=No Information=-"
        lblsalary.Caption = "-=No Information=-"
        MsgBox "No record associated with this id has been found.", vbApplicationModal + vbExclamation, "Not Found"
    End If
    Set rs = Nothing
    Timer2.Enabled = False
    lbldone.Caption = "0 % Completed"
    lblsearch.Caption = "Search Status"
    lblwait.Caption = "Stopped"
    Timer1.Enabled = False
    p1.Value = 0
    txtsearch.Enabled = True
    txtsearch.Text = ""
    txtsearch.SetFocus
    cmdfind.Enabled = True
End If
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdfind_Click
End Sub
