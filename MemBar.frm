VERSION 5.00
Begin VB.Form MemBar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   3765
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   3300
   Icon            =   "MemBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHook 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      ScaleHeight     =   225
      ScaleWidth      =   1425
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2160
      Top             =   1440
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   6
      Left            =   2520
      Picture         =   "MemBar.frx":0442
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   5
      Left            =   2160
      Picture         =   "MemBar.frx":0884
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   4
      Left            =   1800
      Picture         =   "MemBar.frx":0CC6
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   3
      Left            =   1440
      Picture         =   "MemBar.frx":1108
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   2
      Left            =   1080
      Picture         =   "MemBar.frx":154A
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "MemBar.frx":198C
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image img 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "MemBar.frx":1DCE
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   11
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   720
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   720
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "MemBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------
'API for making the form stays on top
'-----------------------------------------------------------------------
Private Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

'-----------------------------------------------------------------------
'API for retrieving memory status
'-----------------------------------------------------------------------
Private Type MEMORYSTATUS
   dwLength As Long
   dwMemoryLoad As Long
  dwTotalPhys As Long
  dwAvailPhys As Long
  dwTotalPageFile As Long
  dwAvailPageFile As Long
  dwTotalVirtual As Long
  dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Dim memInfo As MEMORYSTATUS
'-----------------------------------------------------------------------
'API for tray icon
'taken from the API Guide (http://www.allapi.net/)
'-----------------------------------------------------------------------
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim TrayI As NOTIFYICONDATA
'-----------------------------------------------------------------------

Private Const NUM_BARS = 4                  'number of status bar rows
Private cpu As Object                       'used to retrieve cpu usage
Private isOnTop As Boolean                  'whether form has been set on top

'This API call sets the bar to the topmost form
Public Sub MakeTopMost(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Private Sub Form_Load()
    isOnTop = False                         'form not on top yet
    Me.Width = 1500                         'default form width
    Me.Height = 1120                        'default form height
    
    TrayI.cbSize = Len(TrayI)
    'Set the window's handle (this will be used to hook the specified window)
    TrayI.hWnd = picHook.hWnd
    'Application-defined identifier of the taskbar icon
    TrayI.uId = 1&
    'Set the flags
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Set the callback message
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    'Set the picture (must be an icon!)
    TrayI.hIcon = Me.Icon
                                            'give our form high thread priority
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    If IsWinNT Then                         'check the OS version
        Set cpu = New clsCPUUsageNT
    Else
        Set cpu = New clsCPUUsage
    End If
    
    Timer1_Timer                            'start the timer
End Sub

'resize all form component when form is resized
Private Sub Form_Resize()
    Dim w, h As Integer
    w = Me.Width
    h = Me.Height
    Dim i As Integer
    For i = 0 To (NUM_BARS * 3 - 1)
        lbl(i).Top = 0
        lbl(i).Left = 0
        If i > 2 Then lbl(i).Top = lbl(i - 3).Top + lbl(i - 3).Height - 8
        lbl(i).Width = w - 120
        lbl(i).Height = (h - 360) / NUM_BARS
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'remove the icon
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = picHook.hWnd
    TrayI.uId = 1&
    Shell_NotifyIcon NIM_DELETE, TrayI
    
    Set MemBar = Nothing
    End
End Sub


'double click to close application
Private Sub lbl_DblClick(Index As Integer)
    Unload Me
End Sub

Private Sub lbl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        'Create the icon
        Shell_NotifyIcon NIM_ADD, TrayI
        Me.WindowState = vbMinimized
        Me.Hide
    End If
End Sub

Private Sub picHook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONUP Or Msg = WM_RBUTTONUP Then
        Shell_NotifyIcon NIM_DELETE, TrayI
        Me.Show
        Me.WindowState = vbNormal
    End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

    If Not isOnTop Then                         'make the form go on top
        MakeTopMost Me.hWnd
        isOnTop = True
    End If
    GlobalMemoryStatus memInfo                  'retrieve memory info structure
    
    Me.Caption = Format(Now(), "ddd, d mmm yy")
                                                'get physical memory
    high = byteToMB(memInfo.dwTotalPhys)
    current = byteToMB(memInfo.dwAvailPhys)
    lbl(1).Width = ((high - current) / high) * lbl(0).Width
    lbl(2).Caption = "Physical (" & Round((high - current) / high * 100, 0) & "%)"
    lbl(2).ToolTipText = "Physical Free: " & current & "MB / " & high & "MB (" & 100 - Round((high - current) / high * 100, 0) & "%)"
   
    'Set the tooltiptext for tray icon
    TrayI.szTip = "Physical Free: " & current & "MB / " & high & "MB (" & 100 - Round((high - current) / high * 100, 0) & "%)" & vbNullChar
    Dim bar As Integer
    bar = Round((high - current) / high * 7, 0) - 1
    Me.Icon = img(bar).Picture
    TrayI.hIcon = Me.Icon
    'Update the icon
    Shell_NotifyIcon NIM_MODIFY, TrayI
   
                                                'get paged memory
    high = byteToMB(memInfo.dwTotalPageFile)
    current = byteToMB(memInfo.dwAvailPageFile)
    lbl(4).Width = ((high - current) / high) * lbl(0).Width
    lbl(5).Caption = "Kernel (" & Round((high - current) / high * 100, 0) & "%)"
    lbl(5).ToolTipText = "Kernel Free: " & current & "MB / " & high & "MB (" & 100 - Round((high - current) / high * 100, 0) & "%)"
    
                                                'get virtual memory
    high = byteToMB(memInfo.dwTotalVirtual)
    current = byteToMB(memInfo.dwAvailVirtual)
    lbl(7).Width = ((high - current) / high) * lbl(0).Width
    lbl(8).Caption = "Virtual (" & Round((high - current) / high * 100, 0) & "%)"
    lbl(8).ToolTipText = "Virtual Free: " & current & "MB / " & high & "MB (" & 100 - Round((high - current) / high * 100, 0) & "%)"
        
    high = 100                                  'get CPU usage
    current = cpu.Query
    lbl(10).Width = ((current) / high) * lbl(0).Width
    lbl(11).Caption = "CPU (" & current & "%)"
    lbl(11).ToolTipText = "System Idle (" & high - current & "%)"
    
    

End Sub

'convert bytes to MB equivalent
Private Function byteToMB(ByVal b As Long)
    byteToMB = Round(((b / 1024) / 1024), 2)
End Function

