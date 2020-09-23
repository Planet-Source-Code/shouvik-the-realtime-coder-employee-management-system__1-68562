VERSION 5.00
Begin VB.Form frmsplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1305
   ClientLeft      =   3660
   ClientTop       =   3345
   ClientWidth     =   4755
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
   Moveable        =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   375
      Top             =   630
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   1080
      Top             =   585
   End
   Begin VB.Label lbltoday 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   195
      Left            =   2100
      TabIndex        =   2
      Top             =   1065
      Width           =   555
   End
   Begin VB.Label lblver 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version : "
      Height          =   195
      Left            =   2085
      TabIndex        =   1
      Top             =   345
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   802
      TabIndex        =   0
      Top             =   105
      Width           =   3150
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_Load()
Left = (Screen.Width - Me.Width) / 2
Top = (Screen.Height - Me.Height) / 2

lblver.Caption = "Version : " & App.Major & "." & App.Minor & "." & App.Revision
Timer1_Timer
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
lbltoday.Caption = "Today : " & Format(Date, "dddd,dd mmmm yyyy") & " " & Time
End Sub

Private Sub Timer2_Timer()
i = i + 10
If i = 100 Then
    Unload Me
    frmlogin.Show
End If
End Sub
