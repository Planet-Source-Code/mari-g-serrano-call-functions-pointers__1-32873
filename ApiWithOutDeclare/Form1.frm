VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Call Pointers"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "RltMoveMemory"
      Height          =   435
      Index           =   3
      Left            =   60
      TabIndex        =   3
      Top             =   1575
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RltMoveMemory"
      Height          =   435
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GetCursorPos && FlashWindow &"
      Height          =   435
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   555
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SetWindowText A/W"
      Height          =   435
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Sub Command1_Click(Index As Integer)
  Dim x As Long
  
  Select Case Index
  Case 0 'TEST SetWndText
    Dim s() As Byte
    Dim s1 As String
    s = StrConv("MaRiØ ANSI" & Chr(0), vbFromUnicode)
    s1 = "MaRIO UNICODE" & Chr(0)
    'Unicode: (not in win95-98)
    x = CallApiByName("user32", "SetWindowTextW", Me.hwnd, StrPtr(s1))
    'ANSI:
    x = CallApiByName("user32", "SetWindowTextA", Command1(0).hwnd, VarPtr(s(0)))
 Case 1 'test FlashWindow & GetCursorPos
    Dim c As POINTAPI
    x = CallApiByName("user32", "GetCursorPos", VarPtr(c))
    MsgBox "Cursor Pos: " & c.x & ";" & c.y
    x = CallApiByName("user32", "FlashWindow", Me.hwnd, 1&)
 Case 2
    Dim a As String, b As String
    a = "I´m a String."
    b = "I´m a New Str"
    MsgBox "a= " & a
    x = CallApiByName("kernel32", "RtlMoveMemory", StrPtr(a), StrPtr(b), LenB(a))
    MsgBox "a= " & a
 
 Case 3 'call ptr
    '    Dim lib As Long
    '    lib = LoadLibrary("user32")
    '    x = GetProcAddress(lib, "FlashWindow")
    '    CallPtr x, Me.hwnd, 1&
    '    FreeLibrary lib
    x = CallPtr(AddressOf Test, 13)
    ' x must be 14...
    Debug.Print "x= " & x
    x = CallPtr(AddressOf TestMsg)
 Case 4 'Call ASM
    If MsgBox("This Function Does not work..." & vbCrLf & "Do you want to see a crash?", vbYesNo, "Call ASM Code") = vbYes Then
        x = CallASM
    End If
 End Select

End Sub
