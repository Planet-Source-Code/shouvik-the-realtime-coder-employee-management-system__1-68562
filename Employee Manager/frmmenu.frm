VERSION 5.00
Begin VB.Form frmmenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Management System"
   ClientHeight    =   2460
   ClientLeft      =   3375
   ClientTop       =   2835
   ClientWidth     =   4470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4470
   Begin VB.CommandButton cmdlaunch 
      Caption         =   "Launch"
      Height          =   255
      Left            =   1830
      TabIndex        =   4
      Top             =   2130
      Width           =   1260
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   3150
      TabIndex        =   5
      Top             =   2130
      Width           =   1260
   End
   Begin VB.OptionButton optsearch 
      BackColor       =   &H00000000&
      Caption         =   "Search For Employee Data"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   780
      TabIndex        =   3
      Top             =   1275
      Width           =   2565
   End
   Begin VB.OptionButton optremove 
      BackColor       =   &H00000000&
      Caption         =   "Remove Existing Employee Data"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   780
      TabIndex        =   2
      Top             =   960
      Width           =   3120
   End
   Begin VB.OptionButton optedit 
      BackColor       =   &H00000000&
      Caption         =   "Modify Existing Employee Data"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   780
      TabIndex        =   1
      Top             =   645
      Width           =   2970
   End
   Begin VB.OptionButton optadd 
      BackColor       =   &H00000000&
      Caption         =   "Add New Employee Data"
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   780
      TabIndex        =   0
      Top             =   330
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   30
      TabIndex        =   6
      Top             =   30
      Width           =   4410
      Begin VB.Label lblmsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select A Service And Click Launch"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   750
         TabIndex        =   7
         Top             =   0
         Width           =   2820
      End
   End
   Begin VB.Label lbldes 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Employee Management System."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   510
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   4350
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      Height          =   2445
      Left            =   15
      Top             =   15
      Width           =   4455
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdexit_Click()
confirm = MsgBox("Are you sure to exit Employee Manager ?", vbApplicationModal + vbYesNo + vbInformation, "Exit Application")
If confirm = vbNo Then
    optadd.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    MsgBox "Thanks for being worked with this application.", vbApplicationModal, "GoodBye"
    End
End If
End Sub

Private Sub cmdexit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldes.Caption = "Exits employee manager application."
End Sub

Private Sub cmdlaunch_Click()
If optadd.Value = False And optedit.Value = False And optremove.Value = False And optsearch.Value = False Then
    MsgBox "You should select an option in order to proceed.", vbApplicationModal + vbExclamation, "Nothing Is Selected"
    optadd.SetFocus
ElseIf optadd.Value = True Then
    frmadditon.Show vbModal
ElseIf optedit.Value = True Then
    frmmodification.Show vbModal
ElseIf optremove.Value = True Then
    frmdeletion.Show vbModal
ElseIf optsearch.Value = True Then
    frmsearching.Show vbModal
End If
End Sub

Private Sub cmdlaunch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldes.Caption = "Starts the service associated based on your selection."
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldes.Caption = "Welcome to Employee Management System"
End Sub

Private Sub Form_Unload(Cancel As Integer)
confirm = MsgBox("Are you sure to exit Employee Manager ?", vbApplicationModal + vbYesNo + vbInformation, "Exit Application")
If confirm = vbNo Then
    Cancel = vbNo
    optadd.SetFocus
    Exit Sub
ElseIf confirm = vbYes Then
    MsgBox "Thanks for being worked with this application.", vbApplicationModal, "GoodBye"
    End
End If
End Sub

Private Sub optadd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldes.Caption = "Opens addition manager and helps you to create new employee record that is going to be stored in employee database."
End Sub

Private Sub optedit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldes.Caption = "Opens modification manager and helps you to edit existing employee record that is going to be updated in employee database."
End Sub

Private Sub optremove_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldes.Caption = "Opens deletion manager and helps you to remove existing employee record that is going to be removed from employee database based on the option you choose."
End Sub

Private Sub optsearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldes.Caption = "Opens searching manager and helps you to find out existing employee record that is going to be fetched from employee database based on the employee id you give."
End Sub
