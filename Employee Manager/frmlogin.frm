VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmlogin 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome To System"
   ClientHeight    =   2340
   ClientLeft      =   3375
   ClientTop       =   3390
   ClientWidth     =   4380
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
   Icon            =   "frmlogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4380
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   3615
      Top             =   1845
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   135
      Left            =   270
      TabIndex        =   9
      Top             =   1890
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   2475
      Top             =   1485
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "Change"
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   1080
      Width           =   1260
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2925
      TabIndex        =   3
      Top             =   1080
      Width           =   1260
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Save And Go"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1260
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1515
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   690
      Width           =   2715
   End
   Begin VB.TextBox txtusername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1515
      TabIndex        =   0
      Top             =   360
      Width           =   2715
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   4335
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please LogIn"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1620
         TabIndex        =   5
         Top             =   0
         Width           =   1080
      End
   End
   Begin VB.Label lbldone 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Completed"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1605
      TabIndex        =   12
      Top             =   2070
      Width           =   1185
   End
   Begin VB.Label lblwait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   3330
      TabIndex        =   11
      Top             =   1575
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading System..."
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   75
      TabIndex        =   10
      Top             =   1575
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   255
      TabIndex        =   7
      Top             =   750
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   420
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   1410
      Left            =   15
      Top             =   15
      Width           =   4380
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub cmdcancel_Click()
confirm = MsgBox("Are you sure to exit ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Exit")
If confirm = vbYes Then
    End
Else
    txtusername.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmdchange_Click()
frmverify.Show vbModal
End Sub

Private Sub cmdok_Click()
If Trim(username) = "" And Trim(password) = "" Then
    If Trim(txtusername.Text) <> "" And Trim(txtpassword.Text) <> "" Then
        SaveSetting "Employee", "Login", "Username", Trim(txtusername.Text)
        SaveSetting "Employee", "Login", "Password", Trim(txtpassword.Text)
        MsgBox "User authentication has been accepted." & vbCrLf & "Click Ok to launch the system.", vbApplicationModal + vbInformation, _
            "User Data Saved"
        lblmsg.Enabled = False
        Label2.Enabled = False
        Label3.Enabled = False
        txtpassword.Enabled = False
        txtusername.Enabled = False
        cmdcancel.Enabled = False
        cmdchange.Enabled = False
        cmdok.Enabled = False
        Me.Height = 2805
        Left = (Screen.Width - Me.Width) / 2
        Top = (Screen.Height - Me.Height) / 2
        Timer1.Enabled = True
        Timer2.Enabled = True
    Else
        MsgBox "Sorry,blank values cannot be accepted in user settings.", vbApplicationModal + vbCritical, "Error Accepting Values"
        txtusername.SetFocus
        Exit Sub
    End If
ElseIf Trim(username) <> "" And Trim(password) <> "" Then
    If Trim(txtusername.Text) = Trim(username) And Trim(txtpassword.Text) = Trim(password) Then
        lblmsg.Enabled = False
        Label2.Enabled = False
        Label3.Enabled = False
        txtpassword.Enabled = False
        txtusername.Enabled = False
        cmdcancel.Enabled = False
        cmdchange.Enabled = False
        cmdok.Enabled = False
        Me.Height = 2805
        Left = (Screen.Width - Me.Width) / 2
        Top = (Screen.Height - Me.Height) / 2
        Timer1.Enabled = True
        Timer2.Enabled = True
    Else
        MsgBox "The system cannot recognised these user settings." & vbCrLf & "Please enter valid data and try again...", vbApplicationModal + _
            vbExclamation, "Invalid User Data"
        txtusername.Text = ""
        txtpassword.Text = ""
        txtusername.SetFocus
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Activate()
username = GetSetting("Employee", "Login", "Username")
password = GetSetting("Employee", "Login", "Password")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If cmdcancel.Enabled = True Then
        cmdcancel_Click
    ElseIf cmdcancel.Enabled = False Then
        Exit Sub
    End If
End If
End Sub

Private Sub Form_Load()
Me.Height = 1905
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2
p1.Max = 100
p1.Min = 0
Timer1.Enabled = False
Timer2.Enabled = False

username = GetSetting("Employee", "Login", "Username")
password = GetSetting("Employee", "Login", "Password")

If Trim(username) = "" And Trim(txtpassword) = "" Then
    MsgBox "First of all thank you for downloading this application." & vbCrLf & "If you find any bugs please notify it by posting your" & vbCrLf & _
        "comment.Any suggestions for improving this application" & vbCrLf & "will be appreciated.", vbApplicationModal, "Thank You Fellow Programmer"
    MsgBox "This application is running first time on this system.please set an" & vbCrLf & "user authentication in order to prevent unauthorized access.", _
            vbApplicationModal + vbInformation, "Welcome"
    cmdok.Caption = "Save And Go"
    cmdchange.Enabled = False
    lblmsg.Caption = "Please Set User Authentication"
Else
    cmdok.Caption = "GO"
    cmdchange.Enabled = True
    lblmsg.Caption = "Please Login"
End If
End Sub

Private Sub Timer1_Timer()
i = i + 1
p1.Value = i
lbldone.Caption = p1.Value & "% Completed"
If i = 100 Then
    Unload Me
    frmmenu.Show
End If
End Sub

Private Sub Timer2_Timer()
lblwait.Visible = Not lblwait.Visible
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdok_Click
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdok_Click
End Sub
