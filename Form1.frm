VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMessage 
      Caption         =   "Messagebox"
      Height          =   360
      Left            =   150
      TabIndex        =   15
      Top             =   195
      Width           =   1455
   End
   Begin VB.CommandButton cmdInvoked 
      Caption         =   "invoke"
      Height          =   360
      Left            =   3930
      TabIndex        =   14
      Top             =   1965
      Width           =   990
   End
   Begin VB.CommandButton cmdIO 
      Caption         =   "IO"
      Height          =   360
      Left            =   1710
      TabIndex        =   13
      Top             =   195
      Width           =   990
   End
   Begin VB.CommandButton cmdUIA 
      Caption         =   "UIA"
      Height          =   360
      Left            =   3990
      TabIndex        =   12
      Top             =   195
      Width           =   990
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process"
      Height          =   360
      Left            =   2790
      TabIndex        =   11
      Top             =   195
      Width           =   1095
   End
   Begin VB.CommandButton cmdApp 
      Caption         =   "App"
      Height          =   360
      Left            =   2655
      TabIndex        =   10
      Top             =   1965
      Width           =   990
   End
   Begin VB.CommandButton cmdMy 
      Caption         =   "My"
      Height          =   360
      Left            =   150
      TabIndex        =   9
      Top             =   795
      Width           =   990
   End
   Begin VB.CommandButton cmdThread 
      Caption         =   "Thread"
      Height          =   360
      Left            =   3030
      TabIndex        =   8
      Top             =   795
      Width           =   990
   End
   Begin VB.CommandButton cmdClipboard 
      Caption         =   "Clipboard"
      Height          =   360
      Left            =   4230
      TabIndex        =   7
      Top             =   795
      Width           =   1215
   End
   Begin VB.CommandButton cmdStopWatch 
      Caption         =   "StopWatch"
      Height          =   360
      Left            =   150
      TabIndex        =   6
      Top             =   1395
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnvironment 
      Caption         =   "Environment"
      Height          =   360
      Left            =   1350
      TabIndex        =   5
      Top             =   795
      Width           =   1590
   End
   Begin VB.CommandButton cmdPorts 
      Caption         =   "Ports"
      Height          =   360
      Left            =   1590
      TabIndex        =   4
      Top             =   1395
      Width           =   990
   End
   Begin VB.CommandButton cmdClock 
      Caption         =   "Clock"
      Height          =   360
      Left            =   2790
      TabIndex        =   3
      Top             =   1395
      Width           =   990
   End
   Begin VB.CommandButton cmdTimer 
      Caption         =   "Timer"
      Height          =   360
      Left            =   3990
      TabIndex        =   2
      Top             =   1395
      Width           =   990
   End
   Begin VB.CommandButton cmdSendkeys 
      Caption         =   "Sendkeys"
      Height          =   360
      Left            =   210
      TabIndex        =   1
      Top             =   1965
      Width           =   990
   End
   Begin VB.CommandButton cmdScreen 
      Caption         =   "Screen"
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   1965
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sw                 As New StopWatch   'instatiate
Private WithEvents tmr As Timer
Attribute tmr.VB_VarHelpID = -1
'TODO
'automation
'sendkeys
'driveinfo class
'etc
'
Private Sub cmdMessage_Click()
    'Display a simple message box with api wrapped in .net-like class collection
    Messagebox.Show "Hello from a messagebox class on the regular desktop.", "", 0, -1
    'Get authorization from the physical user. For some button feature via automation or general question
    Dim mb As MB_RESULT
    mb = Messagebox.Authorize("Question that you authorize on a secure desktop?", "Secured message", 5000) 'timout in 5 seconds
    If mb = IYES Then
      MsgBox "Yes"
    ElseIf mb = INO Then
      MsgBox "No"
    Else
      MsgBox "Not answered"
    End If
End Sub

Private Sub cmdIO_Click()
    Directory.Create ("C:\tinkle") 'Create a test directory
    '
    'If folder was created successfully
    If Directory.Exists("C:\tinkle\") = True Then
        'Create a file in the newly created folder
        File.Create ("C:\tinkle\test.txt")
        '
        'Check to see if the new path exists
        If File.Exists("C:\tinkle\test.txt") = True Then
            Messagebox.Show "File created"
            File.Delete ("C:\tinkle\test.txt")
            If File.Exists("C:\tinkle\test.txt") = False Then
                Messagebox.Show "File deleted"
            Else
                Messagebox.Show "Error deleting file"
            End If
            Directory.Delete ("C:\tinkle\")
            If Directory.Exists("C:\tinkle\") = False Then
                Messagebox.Show "Directory deleted"
            Else
                Messagebox.Show "Error deleting Directory"
            End If
        End If
    End If
End Sub

Private Sub cmdMy_Click()
  MsgBox "Tick count: " & My.Computer.Clock.TickCount
  
End Sub

Private Sub cmdProcess_Click()
    '
    'Get process objects running on the computer
    Dim p() As Process
    p = System.Diagnostics.Process.GetProcesses
    Messagebox.Show UBound(p) & " processes running"
      Dim txt As String
      Dim i As Long
      Do
        txt = txt & p(i).ProcessName & vbCrLf
        If i = UBound(p) Then Exit Do
        i = i + 1
      Loop
      Messagebox.Show txt & vbCrLf & "Starting notepad..."
    '
    'Start a process without arguments
    Call Process.Start("C:\Windows\Notepad.exe")
End Sub
Private Sub cmdUIA_Click()
    'Invoke a button click with UIA (user interface automation)
    Call System.Windows.Automation.InvokeElement(Me.hwnd, "invoke")
    '
    '
End Sub
Private Sub cmdApp_Click()
    Application.DoEvents_
    Messagebox.Show "events done"
    '
    'Get the path that this application started from
    Messagebox.Show "app started here: " & Application.StartupPath
End Sub
Private Sub cmdEnvironment_Click()
    'Get a version object for the current operating system
    Messagebox.Show "Operating system Minor version: " & Environment.OSVersion.Version.dwMinorVersion
    '
End Sub
Private Sub cmdThread_Click()
    'Sleep for 1/5 a second
    Dim ret As Long
    ret = Threading.Thread.Sleep(200)
    Messagebox.Show "Slept 200 milliseconds " & CBool(ret)
    '
End Sub
Private Sub cmdClipboard_Click()
    'Get text from the clipboard
    Messagebox.Show (Clipboard.GetText)
    '
End Sub
Private Sub cmdStopWatch_Click()
    'Initialize a new stopwatch timer and start it
    Dim hwnd As Long
    hwnd = sw.StartNew 'initialize
    sw.Start (hwnd) 'start
    Threading.Thread.Sleep (1200) 'Sleep for about a second
    Dim Ticks   As Double
    Dim Seconds As Long
    Ticks = sw.ElapsedTicks(hwnd) 'measure time interval in ticks (the smallest increment)
    Seconds = sw.ElapsedMilliseconds(hwnd) 'return normal millsecond value
    Messagebox.Show Ticks & vbCrLf & "This value should be near 1.200" & vbCrLf & vbCrLf & Seconds & vbCrLf & "This value should be near 1200"
    '
    '
End Sub
Private Sub cmdInvoked_Click()
  
  Messagebox.Show "invoked button click from automation"
End Sub
Private Sub cmdPorts_Click()
    ' Get port names in specified server
    Dim port() As SerialPort_
    port = IO.Ports.SerialPort.GetSerialPorts("")
    Dim i As Long
    For i = 0 To UBound(port) - 1
        With port(i)
            Messagebox.Show "Port Name: " & .pPortName & vbCrLf & "Description: " & .pDescription & vbCrLf & "Port type: " & .fPortType & vbCrLf & "Monitor name: " & .pMonitorName
        End With
    Next
    '    'Get serial ports on specified server
    '    Dim portnames() As String
    '    portnames = System.IO.Ports.SerialPort.GetPortNames("")
    '    Dim i As Long
    '    For i = 0 To UBound(portnames) - 1
    '        MessageBox.Show portnames(i)
    '    Next
End Sub
Private Sub cmdClock_Click()
    MsgBox My.Computer.Clock.localtime.wMinute
End Sub
Private Sub cmdTimer_Click()
    Set tmr = New Timer
    tmr.Interval = 400
    tmr.Enabled = True
End Sub
Private Sub Command1_Click()
    Dim t As New TimeSpan
End Sub
Private Sub tmr_Timer()
    tmr.Enabled = False
    tmr.Interval = 0
    Messagebox.Show "Timer has elapsed"
End Sub

Private Sub cmdSendkeys_Click()
    Dim disp As Long
    
    disp = SendKeys.Flush
    Messagebox.Show "flushed/dispatched " & disp & " messages"
End Sub
Private Sub cmdScreen_Click()
        On Error Resume Next
        Dim h   As Long
        Dim s() As Screen
        s = My.Computer.Screen.AllScreens
        Dim i As Long
        Dim txt As String
        For i = 0 To UBound(s)
            txt = txt & "Primary display: " & s(i).Primary & vbCrLf
            txt = txt & "WidthxHeight: " & s(i).Bounds.Size.Width_ & "x" & s(i).Bounds.Size.Height_ & vbCrLf
            txt = txt & "Monitor Handle: " & s(i).Handle & vbCrLf
            txt = txt & "Working area size WidthXHeight: " & s(i).WorkingArea.Size.Width_ & "x" & s(i).WorkingArea.Size.Height_
            txt = txt & vbCrLf & vbCrLf
        Next
        MsgBox txt
        '
        '
        'Get device names
        Dim n   As Long
        Dim d() As DISPLAY_DEVICE
        d = My.Computer.Screen.GetDisplayDevices
        For n = 0 To UBound(d)
            MsgBox "Device name: " & d(n).DeviceName
        Next
End Sub

