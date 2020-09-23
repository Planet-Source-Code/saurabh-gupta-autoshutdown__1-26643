VERSION 5.00
Begin VB.Form Shutdown 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AutoShutdown"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   Icon            =   "Shutdown.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSeconds 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtMinutes 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtHours 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   600
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Countdown"
      Height          =   375
      Left            =   968
      TabIndex        =   0
      Tag             =   "0"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   960
   End
   Begin VB.Label Label4 
      Caption         =   "Shutdown in:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Seconds:"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Minutes:"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Hours:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   495
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show Window"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Shutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Dim nid As NOTIFYICONDATA

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.
'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" _
    Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Counter As Long
Private ShutdownTime As Long

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Sub Command1_Click()
    If Command1.Tag = 0 Then
        StartCounter
    Else
        ResetCounter
    End If
End Sub


Private Sub Timer1_Timer()
    txtSeconds.Text = txtSeconds.Text - 1
    If txtSeconds.Text = 0 And txtMinutes.Text = 0 And txtHours.Text = 0 Then
        ExitWindowsEx 1, 0
        Unload Me
    End If
End Sub

Private Sub txtHours_Validate(KeepFocus As Boolean)
    If Not IsNumeric(txtHours.Text) Then
        txtHours.Text = "0"
        KeepFocus = True
    End If
End Sub

Private Sub txtMinutes_Validate(KeepFocus As Boolean)
    If Not IsNumeric(txtMinutes.Text) Then
        txtMinutes.Text = "0"
        KeepFocus = True
    End If
End Sub

Private Sub txtMinutes_Change()
    If Not IsNumeric(txtMinutes.Text) Then
        Exit Sub
    End If
    If txtMinutes.Text > 59 Then
        txtMinutes.Text = txtMinutes.Text Mod 60
        If txtHours.Text < 10 Then
            txtHours.Text = txtHours.Text + 1
        End If
    End If
    If txtMinutes.Text < 0 Then
        txtMinutes.Text = 59
        txtHours.Text = txtHours.Text - 1
    End If
End Sub

Private Sub txtSeconds_Validate(KeepFocus As Boolean)
    If Not IsNumeric(txtSeconds.Text) Then
        txtSeconds.Text = "0"
        KeepFocus = True
    End If
End Sub

Private Sub txtSeconds_Change()
    If Not IsNumeric(txtSeconds.Text) Then
        Exit Sub
    End If
    If txtSeconds.Text > 59 Then
        txtSeconds.Text = txtSeconds.Text Mod 60
        txtMinutes.Text = txtMinutes.Text + 1
    End If
    If txtSeconds.Text < 0 Then
        txtSeconds.Text = 59
        txtMinutes.Text = txtMinutes.Text - 1
    End If
End Sub
Private Sub StartCounter()
    ShutdownTime = txtHours.Text + txtMinutes.Text + txtSeconds.Text
    If ShutdownTime = 0 Then
        Exit Sub
    End If
    Command1.Caption = "Stop Countdown"
    Command1.Tag = 1
    Counter = 0
    txtHours.Enabled = False
    txtMinutes.Enabled = False
    txtSeconds.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub ResetCounter()
    Command1.Caption = "Start Countdown"
    ShutdownTime = 0
    Command1.Tag = 0
    Timer1.Enabled = False
    txtHours.Enabled = True
    txtMinutes.Enabled = True
    txtSeconds.Enabled = True
End Sub

Private Sub setSysTrayIcon()
    nid.cbSize = Len(nid)
    nid.hwnd = Me.hwnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = Me.Icon
    nid.szTip = "AutoShutdown" & vbNullChar

    'Call the Shell_NotifyIcon function to add the icon to the taskbar
    'status area.
    Shell_NotifyIcon NIM_ADD, nid
End Sub
Private Sub Form_Load()
    If App.PrevInstance Then
        End
    End If
    setSysTrayIcon
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Me.Visible = False
        Cancel = True
    End If
End Sub
Private Sub Form_Terminate()
    Shell_NotifyIcon NIM_DELETE, nid
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim msg As Long
    Dim sFilter As String
    msg = X / Screen.TwipsPerPixelX
    Select Case msg
        Case WM_LBUTTONDOWN
            If Me.Visible = True Then
                Me.SetFocus
            End If
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            If Me.Visible = False Then
                Me.Visible = True
                Me.SetFocus
            End If
        Case WM_RBUTTONDOWN
            PopupMenu mnuPopup
        Case WM_RBUTTONUP
        Case WM_RBUTTONDBLCLK
    End Select
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub
Private Sub mnuAbout_Click()
    MsgBox "Programmed By Saurabh" + vbCrLf + "http://saurabhonline.8m.net"
End Sub
Private Sub mnuShow_Click()
    If Me.Visible = False Then
        Me.Visible = True
        Me.SetFocus
    End If
End Sub
