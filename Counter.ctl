VERSION 5.00
Begin VB.UserControl Counter 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   ForwardFocus    =   -1  'True
   PropertyPages   =   "Counter.ctx":0000
   ScaleHeight     =   136
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   155
   ToolboxBitmap   =   "Counter.ctx":0032
   Begin VB.PictureBox pcSource 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   420
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox pcDigit 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   0
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Rolling Counter"
      Top             =   0
      Width           =   165
   End
End
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
DefLng A-Z
'Property Variables
Private myValue          As Currency
Private myDigits         As Long
Private myExtraX         As Long
Private myExtraY         As Long
Private myPosnX          As Long
Private myPosnY          As Long
'Working Variables
Private IntValue         As Currency
Private PreviousValue    As Currency
Private ThisValue        As Currency
Private Overflow         As Currency 'value where overflow occurs
Private Delta            As Currency
Private MinDelta         As Currency
Private Digit            As Long
Private LenPres          As Long
Private Roll             As Currency
Private BoxWidth         As Long
Private BoxHeight        As Long
Private CharWidth        As Long
Private Recur            As Long 'Control_Resize Recursion Depth
Private i
'Events
Public Event ReachedZero()
Public Event Reached100()
Public Event MouseMove()
Private Declare Sub BitBlt Lib "gdi32" (ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcDC As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal OpCode As Long)

Public Property Let BackColor(ByVal nwBackColor As OLE_COLOR)
Attribute BackColor.VB_Description = "Sets / Returns the Control's BackColor."
Attribute BackColor.VB_HelpID = 10000
Attribute BackColor.VB_UserMemId = -501

    pcSource.BackColor = nwBackColor
    For i = 0 To myDigits - 1
        pcDigit(i).BackColor = nwBackColor
    Next i
    Set Font = pcSource.Font 'repaint pcSource
    PropertyChanged "BackColor"

End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = pcSource.BackColor

End Property

Public Property Get CharacterExtraX() As Long
Attribute CharacterExtraX.VB_Description = "Sets / Returns extra horizontal spacing for each digit."
Attribute CharacterExtraX.VB_HelpID = 10001

    CharacterExtraX = myExtraX

End Property

Public Property Let CharacterExtraX(ByVal nwExtra As Long)

    If nwExtra < (2 - BoxWidth) Or nwExtra > 30 Then
        Err.Raise 380
      Else 'NOT NWEXTRA...
        myExtraX = nwExtra
        Set Font = pcSource.Font
        PropertyChanged "CaracterExtraX"
    End If

End Property

Public Property Get CharacterExtraY() As Long
Attribute CharacterExtraY.VB_Description = "Sets / Returns extra vertical spacing for each digit."
Attribute CharacterExtraY.VB_HelpID = 10002

    CharacterExtraY = myExtraY

End Property

Public Property Let CharacterExtraY(ByVal nwExtra As Long)

    If nwExtra < (2 - BoxHeight) Or nwExtra > 30 Then
        Err.Raise 380
      Else 'NOT NWEXTRA...
        myExtraY = nwExtra
        Set Font = pcSource.Font
        PropertyChanged "CaracterExtraY"
    End If

End Property

Public Property Get ControlName() As String
Attribute ControlName.VB_Description = "Returns the real name of the Control."
Attribute ControlName.VB_HelpID = 10003

  Dim CntrlName As String

    CntrlName = Parent.ActiveControl.Name
    i = Parent.ActiveControl.Index
    If i >= 0 Then
        CntrlName = CntrlName & "(" & Format$(i) & ")"
    End If
    ControlName = CntrlName

End Property

Public Property Get Digits() As Long
Attribute Digits.VB_Description = "Sets / Returns the Control's number of digits."
Attribute Digits.VB_HelpID = 10004
Attribute Digits.VB_MemberFlags = "200"

    Digits = myDigits

End Property

Public Property Let Digits(ByVal nwDigits As Long)

    If nwDigits = 0 Or nwDigits > 9 Then
        Err.Raise 380
      Else 'NOT NWDIGITS...
        Select Case nwDigits
          Case Is > myDigits
            For i = myDigits To nwDigits - 1
                Load pcDigit(i)
                With pcDigit(i)
                    .Top = pcDigit(i - 1).Top
                    .Left = pcDigit(i - 1).Left + BoxWidth - 1
                    .Visible = True
                End With 'PCDIGIT(I)
            Next i
          Case Is < myDigits
            For i = myDigits To nwDigits + 1 Step -1
                Unload pcDigit(i - 1)
            Next i
        End Select
        myDigits = nwDigits
        PropertyChanged "Digits"
        UserControl_Resize
    End If

End Property

Private Sub Display()

    Delta = Abs(myValue - PreviousValue)
    If Delta >= MinDelta Then
        ThisValue = myValue
        If ThisValue < 0 Then
            ThisValue = ThisValue + Overflow
        End If
        IntValue = Int(ThisValue)
        Roll = ThisValue - IntValue
        Digit = IntValue Mod 10
        i = myDigits - 1
        With pcDigit(i)
            BitBlt .hDC, myExtraX \ 2 + myPosnX, myPosnY, CharWidth, BoxHeight, pcSource.hDC, 0, (Digit + Roll) * BoxHeight, vbSrcCopy
            .Refresh
        End With 'PCDIGIT(I)
        For i = myDigits - 2 To 0 Step -1
            If Digit <> 9 Then
                Roll = 0
            End If
            IntValue = IntValue \ 10
            Digit = IntValue Mod 10
            With pcDigit(i)
                BitBlt .hDC, myExtraX \ 2 + myPosnX, myPosnY, CharWidth, BoxHeight, pcSource.hDC, 0, (Digit + Roll) * BoxHeight, vbSrcCopy
                .Refresh
            End With 'PCDIGIT(I)
        Next i
        Select Case True
          Case (PreviousValue < 0 And myValue >= 0) Or (PreviousValue > 0 And myValue <= 0)
            RaiseEvent ReachedZero
          Case (PreviousValue < 100 And myValue >= 100) Or (PreviousValue > 100 And myValue <= 100)
            RaiseEvent Reached100
        End Select
        pcDigit(0).PSet (1, 1), IIf(myValue < 0, pcSource.ForeColor, pcSource.BackColor)
        PreviousValue = myValue
    End If

End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Sets / Returns the font for the Control."
Attribute Font.VB_HelpID = 10005
Attribute Font.VB_UserMemId = -512

    Set Font = pcSource.Font

End Property

Public Property Set Font(ByVal nwFont As Font)

  Dim Dgt As String * 1

    With pcSource
        Set .Font = nwFont
        BoxWidth = 0
        For i = 0 To 9 'find widest Char
            CharWidth = .TextWidth(Format$(i))
            If CharWidth > BoxWidth Then
                BoxWidth = CharWidth
            End If
        Next i
        .Width = BoxWidth
        BoxWidth = BoxWidth + myExtraX + 3 '1 pixel each side plus 1 border
        BoxHeight = .TextHeight("0") + myExtraY
        .Height = BoxHeight * 11 '0 1 2 3 4 5 6 7 8 9 0
        .Cls
        .CurrentY = (myExtraY / 2) - 1 'start for vertical
        For i = 0 To 10
            Dgt = Right$(Format$(i), 1)
            .CurrentX = (CharWidth - .TextWidth(Dgt)) / 2 'to place Char in the middle
            pcSource.Print Dgt '.Print is not exposed by 'With pcSource' (funny, ain't it)
            .CurrentY = .CurrentY + myExtraY 'vertical spacing
        Next i
    End With 'PCSOURCE
    For i = 0 To myDigits - 1 'prepare pcDigit's
        With pcDigit(i)
            .Width = BoxWidth
            .Height = BoxHeight
            .Cls
            If i = 0 Then
                .Left = 0
              Else 'NOT I...
                .Left = (pcDigit(i - 1).Left + BoxWidth - 1)
            End If
        End With 'PCDIGIT(I)
    Next i
    MinDelta = 0.5 / BoxHeight 'skips Display if Delta is less
    PropertyChanged "Font"
    Digits = myDigits 'repaint Control

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Sets / Returns the Control's ForeColor."
Attribute ForeColor.VB_HelpID = 10006
Attribute ForeColor.VB_UserMemId = -513

    ForeColor = pcSource.ForeColor

End Property

Public Property Let ForeColor(ByVal nwForeColor As OLE_COLOR)

    pcSource.ForeColor = nwForeColor
    Set Font = pcSource.Font 'repaint pcSource
    PropertyChanged "ForeColor"

End Property

Private Sub pcDigit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent MouseMove

End Sub

Public Property Let PosnX(ByVal nwPosn As Long)
Attribute PosnX.VB_Description = "Sets / Returns the horizontal placement of each digit in its box."
Attribute PosnX.VB_HelpID = 10007

    If nwPosn < -16 Or nwPosn > 16 Then
        Err.Raise 380
      Else 'NOT NWPOSN...
        myPosnX = nwPosn
        Set Font = pcSource.Font
        PropertyChanged "PosnX"
    End If

End Property

Public Property Get PosnX() As Long

    PosnX = myPosnX

End Property

Public Property Let PosnY(ByVal nwPosn As Long)
Attribute PosnY.VB_Description = "Sets / Returns the vertical placement of each digit in its box."
Attribute PosnY.VB_HelpID = 10008

    If nwPosn < -20 Or nwPosn > 20 Then
        Err.Raise 380
      Else 'NOT NWPOSN...
        myPosnY = nwPosn
        Set Font = pcSource.Font
        PropertyChanged "PosnY"
    End If

End Property

Public Property Get PosnY() As Long

    PosnY = myPosnY

End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Displays the accurate value."
Attribute Refresh.VB_HelpID = 10011

    PreviousValue = Overflow
    Display

End Sub

Private Sub UserControl_Initialize()

    myDigits = 1
    BoxWidth = 1

End Sub

Private Sub UserControl_InitProperties()

    myExtraX = 6
    myExtraY = 6
    myValue = 0
    Set Font = Ambient.Font
    Digits = 3

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        pcSource.BackColor = .ReadProperty("BackColor", &HFFFFFF)
        pcDigit(0).BackColor = pcSource.BackColor
        pcSource.ForeColor = .ReadProperty("ForeColor", &H0&)
        myExtraX = .ReadProperty("CharacterExtraX", 6)
        myPosnX = .ReadProperty("PosnX", 0)
        myExtraY = .ReadProperty("CharacterExtraY", 6)
        myPosnY = .ReadProperty("PosnY", 0)
        myValue = .ReadProperty("Value", 0)
        Set Font = .ReadProperty("Font", Ambient.Font)
        Digits = .ReadProperty("Digits", 3)
    End With 'PROPBAG

End Sub

Private Sub UserControl_Resize()

    Recur = Recur + 1
    Size (BoxWidth - 1) * myDigits * 15 + 15, BoxHeight * 15
    Recur = Recur - 1
    If Recur = 0 Then
        Overflow = 10 ^ myDigits
        PreviousValue = -Overflow 'force repaint
        Display 'repaint Display Value
    End If

End Sub

Private Sub UserControl_Terminate()

    For i = myDigits - 1 To 1 Step -1
        Unload pcDigit(i)
    Next i

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "BackColor", pcSource.BackColor, &HFFFFFF
        .WriteProperty "ForeColor", pcSource.ForeColor, &H0&
        .WriteProperty "CharacterExtraX", myExtraX, 6
        .WriteProperty "PosnX", myPosnX, 0
        .WriteProperty "CharacterExtraY", myExtraY, 6
        .WriteProperty "PosnY", myPosnY, 0
        .WriteProperty "Value", myValue, 0
        .WriteProperty "Font", pcSource.Font, Ambient.Font
        .WriteProperty "Digits", myDigits, 3
    End With 'PROPBAG

End Sub

Public Property Get Value() As Currency
Attribute Value.VB_Description = "Sets / Returns the displayed value."
Attribute Value.VB_HelpID = 10012
Attribute Value.VB_UserMemId = 0

    Value = myValue

End Property

Public Property Let Value(ByVal nwValue As Currency)

    If nwValue > 2147483647 Or myValue < -2147483647 Then
        Err.Raise 380
      Else 'NOT NWVALUE...
        myValue = nwValue
        PropertyChanged "Value"
        Display 'repaint Display Value
    End If

End Property

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-28 11:28)  Decl: 29  Code: 343  Total: 372 Lines
':) CommentOnly: 5 (1,3%)  Commented: 31 (8,3%)  Empty: 87 (23,4%)  Max Logic Depth: 5
