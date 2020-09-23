VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<<File Time Changer>>"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "frmMain.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrow 
      BackColor       =   &H00C0FFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click here to select a file..."
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Select a file by clicking the button on the right..."
      ToolTipText     =   "The name of the file selected to modify the time..."
      Top             =   720
      Width           =   5895
   End
   Begin MSComDlg.CommonDialog CDBox 
      Left            =   120
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   2880
         Picture         =   "frmMain.frx":12CC
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   20
         Top             =   960
         Width           =   360
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   120
         Picture         =   "frmMain.frx":1701
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   19
         Top             =   1030
         Width           =   360
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   120
         Picture         =   "frmMain.frx":1B26
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   18
         Top             =   200
         Width           =   360
      End
      Begin VB.ComboBox lstS 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox lstM 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox lstH 
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdch 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Change it!"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         Picture         =   "frmMain.frx":1F2B
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Click here to make the changes... Not this cannot be undone!"
         Top             =   2280
         Width           =   3495
      End
      Begin VB.TextBox txtt 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "00:00:00"
         ToolTipText     =   "Time that the selected file will be changed into..."
         Top             =   1800
         Width           =   3495
      End
      Begin MSComCtl2.MonthView MonthView1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd.MM.yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   2370
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Click to select the date for the file"
         Top             =   1320
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   1
         MonthBackColor  =   8438015
         ScrollRate      =   1
         StartOfWeek     =   16187394
         TitleBackColor  =   16744576
         TitleForeColor  =   65535
         TrailingForeColor=   12632256
         CurrentDate     =   41033
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(1) Select File (2) choose a date (3) , then time (hr., min., and secs. Once done, click the above button and you are done!"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   3000
         TabIndex        =   15
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Seconds"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5160
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Minutes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4080
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Hours"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3000
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Set the new date for the file"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Select a file to open :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   " and ABSOLUTE NO WARRANTIES!!! USE AT YOUR OWN RISK!!!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   5175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":292D
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public newdate As Date
Public newmnth As Integer
Public newdt As Integer
Public newyr As Integer

Public isdatesel As Boolean
Public isfilesel As Boolean
Public ishrsel As Boolean
Public isminsel As Boolean
Public issecsel As Boolean

Public newhr As Integer
Public newmin As Integer
Public newsec As Integer

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
    Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long


Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long


Private Sub cmdBrow_Click()
'Set the dialog's title
    CD.DialogTitle = "Choose a file ..."
    'Set the dialog's filter
    CD.Filter = "All Files (*.*)|*.*"
    'Show the 'Open File'-dialog
    CD.ShowOpen
    Text1.Text = CD.FileName
    If Text1.Text = "" Then
        isfilesel = False
    Else
        isfilesel = True
    End If
End Sub

Private Sub cmdch_Click()
    Dim m_Date As Date, lngHandle As Long
    Dim udtFileTime As FILETIME
    Dim udtLocalTime As FILETIME
    Dim udtSystemTime As SYSTEMTIME

    If isfilesel = False Then
        MsgBox "You need to select a file first!!!", vbCritical + vbOKOnly, "Error - No file selected"
        Exit Sub
    End If
    
    If isdatesel = False Then
        MsgBox "You need to select a date first!!!", vbCritical + vbOKOnly, "Error - No date selected"
        Exit Sub
    End If
    
    udtSystemTime.wYear = Val(newyr) 'Year(m_Date)
    udtSystemTime.wMonth = Val(newdt) 'dont change this
    udtSystemTime.wDay = Val(newmnth) 'dont change this
    udtSystemTime.wDayOfWeek = MonthView1.DayOfWeek 'Weekday(m_Date) - 1
    udtSystemTime.wHour = Val(newhr) ' Hour(m_Date)
    udtSystemTime.wMinute = Val(newmin) 'Minute(m_Date)
    udtSystemTime.wSecond = Val(newsec) 'Second(m_Date)
    udtSystemTime.wMilliseconds = 0

    ' convert system time to local time
    SystemTimeToFileTime udtSystemTime, udtLocalTime
    ' convert local time to GMT
    LocalFileTimeToFileTime udtLocalTime, udtFileTime
    ' open the file to get the filehandle
    lngHandle = CreateFile(CDBox.FileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    ' change date/time property of the file
    SetFileTime lngHandle, udtFileTime, udtFileTime, udtFileTime
    ' close the handle
    CloseHandle lngHandle

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Call loadhr
Call loadmin
Call loadsec
isdatesel = False
isfilesel = False
ishrsel = False
isminsel = False
issecsel = False

End Sub

Public Function loadhr()
For X = 0 To 23
    lstH.AddItem Format(Str(X), "0#")
Next
lstH.ListIndex = 1
End Function

Public Function loadmin()
For X = 0 To 59
    lstM.AddItem Format(Str(X), "0#")
Next
lstM.ListIndex = 1
End Function

Public Function loadsec()
For X = 0 To 59
    lstS.AddItem Format(Str(X), "0#")
Next
lstS.ListIndex = 1
End Function

Private Sub Label1_Click()
Text1.SetFocus
End Sub

Private Sub Label2_Click()
MonthView1.SetFocus
End Sub

Private Sub Label3_Click()
lstH.SetFocus
End Sub

Private Sub Label4_Click()
lstM.SetFocus
End Sub

Private Sub Label5_Click()
lstS.SetFocus
End Sub

Private Sub lstH_Click()
newhr = lstH.List(lstH.ListIndex)
txtt.Text = lstH.Text & ":" & lstM.Text & ":" & lstS.Text
ishrsel = True
End Sub

Private Sub lstM_Click()
newmin = lstM.List(lstM.ListIndex)
txtt.Text = lstH.Text & ":" & lstM.Text & ":" & lstS.Text
isminsel = True
End Sub

Private Sub lstS_Click()
newsec = lstS.List(lstS.ListIndex)
txtt.Text = lstH.Text & ":" & lstM.Text & ":" & lstS.Text
issecsel = True
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim thisdate() As String
newdate = DateClicked

thisdate() = Split(newdate, "/")
newdt = thisdate(0)
newmnth = thisdate(1)
newyr = thisdate(2)
isdatesel = True
End Sub

Private Sub Text1_GotFocus()
    'retrieve the window which has the focus
    h& = GetFocus&()
    'Create a new cursor
    Call CreateCaret(h&, 1, 10, 25)
    'Show the new cursor
    X& = ShowCaret&(h&)
End Sub

Private Sub txtt_gotfocus()
h& = GetFocus&()
    'Create a new cursor
    Call CreateCaret(h&, 1, 10, 25)
    'Show the new cursor
    X& = ShowCaret&(h&)
    txtt.ShowWhatsThis
End Sub
