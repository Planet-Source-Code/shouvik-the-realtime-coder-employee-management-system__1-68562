VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdelprogress 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deleting..."
   ClientHeight    =   690
   ClientLeft      =   3165
   ClientTop       =   3795
   ClientWidth     =   4050
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   4050
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1125
      Top             =   90
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   135
      Left            =   105
      TabIndex        =   0
      Top             =   300
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   238
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblprogress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Removed"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   45
      Width           =   810
   End
   Begin VB.Label lblwait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   2940
      TabIndex        =   2
      Top             =   45
      Width           =   990
   End
   Begin VB.Label lbldone 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Completed"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   450
      Width           =   1185
   End
End
Attribute VB_Name = "frmdelprogress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer, tot As Integer

Private Sub Form_Activate()
Do Until rs.EOF
    rs.Delete
    i = i + 1
    p1.Value = i
    lblprogress.Caption = "Removed : " & i & "\" & tot
    lbldone.Caption = CInt((i * 100) / tot) & " % Completed"
    rs.MoveNext
Loop
Timer1.Enabled = False
MsgBox "Successfully removed all records.", vbApplicationModal + vbInformation, "Done"
Unload Me
Unload frmdeletion
frmmenu.optadd.SetFocus
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2

i = 0
Set db = OpenDatabase(App.Path & "\employee.mdb")
Set rs = db.OpenRecordset("info", dbOpenTable)
If rs.RecordCount > 0 Then rs.MoveFirst

tot = rs.RecordCount
p1.Max = rs.RecordCount
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
lblwait.Visible = Not lblwait.Visible
End Sub
