VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   360
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1395
   ControlBox      =   0   'False
   FillColor       =   &H00C0FFC0&
   ForeColor       =   &H00C0FFC0&
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   1395
   Begin VB.Timer tmrTick 
      Interval        =   40
      Left            =   0
      Top             =   300
   End
   Begin prjClock.Counter cntTime 
      Height          =   390
      Index           =   0
      Left            =   1005
      TabIndex        =   0
      ToolTipText     =   "Seconds"
      Top             =   -15
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   688
      ForeColor       =   16777215
      CharacterExtraX =   0
      CharacterExtraY =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Digits          =   2
   End
   Begin prjClock.Counter cntTime 
      Height          =   390
      Index           =   1
      Left            =   495
      TabIndex        =   1
      ToolTipText     =   "Minutes"
      Top             =   -15
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   688
      ForeColor       =   16777215
      CharacterExtraX =   0
      CharacterExtraY =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Digits          =   2
   End
   Begin prjClock.Counter cntTime 
      Height          =   390
      Index           =   2
      Left            =   -15
      TabIndex        =   2
      ToolTipText     =   "Hours"
      Top             =   -15
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   688
      ForeColor       =   16777215
      CharacterExtraX =   0
      CharacterExtraY =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Digits          =   2
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Index           =   1
      Left            =   885
      TabIndex        =   4
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Index           =   0
      Left            =   375
      TabIndex        =   3
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left        As Long
    Top         As Long
    Right       As Long
    Bottom      As Long
End Type

Private WinRect             As RECT

Private WithEvents Systray  As clsSystray
Attribute Systray.VB_VarHelpID = -1

Private ShowAgain           As Single
Private SitsOnTop           As Boolean

Private Sub cntTime_MouseMove(Index As Integer)

    If SitsOnTop Then
        Hide
        ShowAgain = (Timer + 5) Mod 68400 'after 5 seconds
    End If

End Sub

Private Sub Form_Load()

  Dim i     As Long
  Dim Delay As Single

    If App.PrevInstance Then
        Unload Me
      Else 'APP.PREVINSTANCE = FALSE/0
        For i = 0 To 2
            cntTime(i).BackColor = BackColor
            cntTime(i).ForeColor = ForeColor
        Next i
        lbl(0).ForeColor = ForeColor
        lbl(1).ForeColor = ForeColor
        Set Systray = New clsSystray
        With Systray
            .SetOwner Me
            .AddIconToTray Icon.Handle, , True
            .Tooltip = vbCrLf & App.ProductName & vbCrLf & vbCrLf & "   Click to unload" & vbCrLf
            .ShowBalloon "        Click me to terminate", "System Tray Clock", InfoIcon
            i = FindWindow("Shell_TrayWnd", vbNullString) 'find tray
            GetWindowRect i, WinRect
            If Not InIDE Then
                SetParent hwnd, i 'tray is gonna be my parent
            End If
            If WinRect.Bottom - WinRect.Top < 64 Then
                Move 0, 30 'on top of start button
                SitsOnTop = True
              Else 'NOT WINRECT.BOTTOM...
                Move 0, 480 'below start button
            End If

            Delay = Timer + 3
            Do
                DoEvents
            Loop Until Timer > Delay
            .HideBalloon
        End With 'SYSTRAY
    End If

End Sub

Private Function InIDE(Optional c As Boolean = False) As Boolean

  Static b  As Boolean

    b = c
    If b = False Then
        Debug.Assert InIDE(True)
    End If
    InIDE = b

End Function

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    cntTime_MouseMove 0

End Sub

Private Sub Systray_MouseDown(Button As Integer)

    Unload Me

End Sub

Private Sub tmrTick_Timer()

  Dim Tim       As Single
  Dim Frac      As Single
  Dim Hour      As Single
  Dim Minute    As Single
  Dim Second    As Single

    Hour = Timer
    If Hour > ShowAgain Then
        Show
        If Not SitsOnTop Then
            ShowAgain = 99999
        End If
    End If
    Tim = Int(Hour)
    Frac = Hour - Tim
    Second = Tim Mod 60 + Frac
    Minute = (Tim \ 60) Mod 60
    Hour = Tim \ 3600
    If Second > 59 Then
        Minute = Minute + Frac
        If Minute > 59 Then
            Hour = Hour + Frac
        End If
    End If
    cntTime(0) = Second
    cntTime(1) = Minute
    cntTime(2) = Hour

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-28 11:28)  Decl: 18  Code: 107  Total: 125 Lines
':) CommentOnly: 2 (1,6%)  Commented: 8 (6,4%)  Empty: 27 (21,6%)  Max Logic Depth: 4
