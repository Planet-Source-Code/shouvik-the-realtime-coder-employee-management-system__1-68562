VERSION 5.00
Begin VB.Form frmverify 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change User Settings"
   ClientHeight    =   3270
   ClientLeft      =   3930
   ClientTop       =   2520
   ClientWidth     =   4335
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
   ScaleHeight     =   3270
   ScaleWidth      =   4335
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2340
      TabIndex        =   4
      Top             =   2925
      Width           =   930
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      Height          =   255
      Left            =   1365
      TabIndex        =   3
      Top             =   2925
      Width           =   930
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3315
      TabIndex        =   5
      Top             =   2925
      Width           =   930
   End
   Begin VB.TextBox txtusername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1545
      TabIndex        =   1
      Top             =   2205
      Width           =   2715
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1545
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2535
      Width           =   2715
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   10
      Top             =   1875
      Width           =   4275
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modify User Data"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1425
         TabIndex        =   11
         Top             =   0
         Width           =   1470
      End
   End
   Begin VB.TextBox txtverify 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   810
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1035
      Width           =   2715
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   7
      Top             =   720
      Width           =   4275
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Current Password"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   855
         TabIndex        =   8
         Top             =   0
         Width           =   2610
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Username : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   13
      Top             =   2265
      Width           =   1395
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password : "
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   2595
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   1410
      Left            =   15
      Top             =   1860
      Width           =   4320
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[ENTER=Accept  ESC=Exit]"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   1065
      TabIndex        =   9
      Top             =   1335
      Width           =   2235
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   930
      Left            =   15
      Top             =   705
      Width           =   4320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only administrator can change current user data. Please verify yourself to the system by entering current password."
      ForeColor       =   &H0000FFFF&
      Height          =   630
      Left            =   105
      TabIndex        =   6
      Top             =   30
      Width           =   4215
   End
End
Attribute VB_Name = "frmverify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
confirm = MsgBox("Are you sure to exit without saving current data ?", vbApplicationModal + vbYesNo + vbQuestion, "Confirm Exit")
If confirm = vbYes Then
    Unload Me
    frmlogin.txtusername.SetFocus
ElseIf confirm = vbNo Then
    txtusername.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmdclear_Click()
txtusername.Text = ""
txtpassword.Text = ""
txtusername.SetFocus
End Sub

Private Sub cmdsave_Click()
If Trim(txtusername.Text) <> "" And Trim(txtpassword.Text) <> "" Then
    SaveSetting "Employee", "Login", "Username", Trim(txtusername.Text)
    SaveSetting "Employee", "Login", "Password", Trim(txtpassword.Text)
    MsgBox "User data have been changed." & vbCrLf & "Use these information during login.", vbApplicationModal + vbInformation, "Done"
    Unload Me
Else
    MsgBox "Sorry,blank values cannot be accepted in user settings.", vbApplicationModal + vbCritical, "Error Accepting Values"
    txtusername.SetFocus
    Exit Sub
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If Label1.Enabled = True Then
        Unload Me
    ElseIf Label1.Enabled = False Then
        cmdcancel_Click
    End If
End If
End Sub

Private Sub Form_Load()
Me.Height = 2115
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdsave_Click
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdsave_Click
End Sub

Private Sub txtverify_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Trim(txtverify.Text) <> "" Then
        If Trim(txtverify.Text) = Trim(password) Then
            Label1.Enabled = False
            Label2.Enabled = False
            Label3.Enabled = False
            txtverify.Enabled = False
            Me.Height = 3750
            Left = (Screen.Width - Me.Width) / 2
            Top = (Screen.Height - Me.Height) / 2
            txtusername.SetFocus
        Else
            MsgBox "The password you typed didn't match with the valid one." & vbCrLf & "Please enter valid password or exit this service.", _
                    vbApplicationModal + vbCritical, "Unauthorized Access"
            txtverify.Text = ""
            txtverify.SetFocus
        End If
    Else
        Beep
        Exit Sub
    End If
End If
End Sub
