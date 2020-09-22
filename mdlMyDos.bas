Attribute VB_Name = "mdlMyDos"
'Filename: mdlMyDos.bas
'Desc    : Simple MS Dos like program
'          This module conatins Main sub and all console functions and constants
'Author  : pramod kumar
'E-mail  :tk_pramod@yahoo.com
'Company : Imprudents (www.imprudents.com)
'Copyright © 2000-2001 Imprudents, All Rights Reserved
'Thanks  :http://www.vb-world.net/articles/console/index.html
'         (Jay Freeman (saurik@saurik.com))
'Date    :11-Feb-2001



Option Explicit
Option Compare Text ' That is, "AAA" is equal to "aaa".

Public Const CON_KEY_EVENT& = &H1
Public Const CON_MOUSE_EVENT& = &H2
Public Const CON_WINDOW_BUFFER_SIZE_EVENT& = &H4
Public Const CON_MENU_EVENT& = &H8
Public Const CON_FOCUS_EVENT& = &H10

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

Public Const FOREGROUND_BLUE& = &H1
Public Const FOREGROUND_GREEN& = &H2
Public Const FOREGROUND_RED& = &H4
Public Const BACKGROUND_BLUE& = &H10
Public Const BACKGROUND_GREEN& = &H20
Public Const BACKGROUND_RED& = &H40

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0

Public Type COORD
  x As Integer
  y As Integer
End Type

Type SMALL_RECT
  left As Integer
  top As Integer
  right As Integer
  bottom As Integer
End Type

Type CONSOLE_SCREEN_BUFFER_INFO
  dwSize As COORD
  dwCursorPosition As COORD
  wAttributes As Integer
  srWindow As SMALL_RECT
  dwMaximumWindowSize As COORD
End Type

Type CONSOLE_CURSOR_INFO
  dwSize As Long
  bVisible As Long
End Type

Private Type WINDOW_BUFFER_SIZE_RECORD
    dwSize As COORD
End Type

Private Type FOCUS_EVENT_RECORD
    bSetFocus As Long
End Type

Private Type MENU_EVENT_RECORD
    dwCommandId As Long
End Type

Private Type MOUSE_EVENT_RECORD
    dwMousePosition As COORD
    dwButtonState As Long
    dwControlKeyState As Long
    dwEventFlags As Long
End Type

Public Type KEY_EVENT_RECORD
    bKeyDown As Long
    wRepeatCount As Integer
    wVirtualKeyCode As Integer
    wVirtualScanCode As Integer
    uChar As Integer
    dwControlKeyState As Long
End Type

Public Type PINPUT_RECORD
    EventType As Integer
    KeyEvent As KEY_EVENT_RECORD
    MouseEvent As MOUSE_EVENT_RECORD
    WindowBufferSizeEvent As WINDOW_BUFFER_SIZE_RECORD
    MenuEvent As MENU_EVENT_RECORD
    FocusEvent As FOCUS_EVENT_RECORD
End Type

Declare Function AllocConsole& Lib "kernel32" ()
Declare Function FreeConsole& Lib "kernel32" ()
Declare Function GetStdHandle& Lib "kernel32" (ByVal nStdHandle As Long)
Declare Function CloseHandle& Lib "kernel32" (ByVal hObject As Long)
Declare Function ReadConsole& Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal _
lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any)

Declare Function ReadConsoleInput% Lib "kernel32" Alias "ReadConsoleInputA" (ByVal hConsoleInput As _
Long, pirBuffer As PINPUT_RECORD, ByVal cInRecords As Long, lpcRead As Long)

Declare Function WriteConsole& Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer _
As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any)

Declare Function FillConsoleOutputAttribute% Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute _
As Long, ByVal nLength As Long, ByVal dwWriteCoord As Long, lpNumberOfAttrsWritten As Long)

Declare Function GetLastError Lib "kernel32" () As Long
Declare Function WaitForSingleObject& Lib "kernel32" (ByVal hObject As Long, ByVal dwTimeout As Long)
Declare Function SetConsoleCtrlHandler& Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long)
Declare Function SetConsoleTitle& Lib "kernel32" Alias "SetConsoleTitleA" (ByVal Title As String)
Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Declare Function SetConsoleTextAttribute% Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttr As Long)

Declare Function WriteConsoleOutputCharacter Lib "kernel32" Alias "WriteConsoleOutputCharacterA" (ByVal _
hConsoleOutput As Long, ByVal lpCharacter As Long, ByVal nLength As Long, ByVal dwWriteCoord As Long, _
ByRef lpNumberOfCharsWritten As Long) As Long

Declare Function WriteConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, lpAttribute _
As Integer, ByVal nLength As Long, ByVal dwWriteCoord As Long, lpNumberOfAttrsWritten As Long) As Long

Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo _
As CONSOLE_SCREEN_BUFFER_INFO) As Long

Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal _
hConsoleOutput As Long, ByVal cCharacter As Long, ByVal nLength As Long, ByVal dwWriteCoord As Long, _
ByRef lpNumberOfCharsWritten As Long) As Long

Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal dwCursorPosition As Long) As Long

Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource _
As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize _
As Long, Arguments As Long) As Long

Public hInConsole As Long
Public hOutConsole As Long
Public sPrompt As String

Public Sub Main()
  Dim Length As Long
  Dim Buffer As String
  Dim myEvent As PINPUT_RECORD
  Dim Quit As Boolean
  sPrompt = ">"
  Buffer = ""
  Quit = False
  
  OpenConsole "My Console App."
  Call SetBkColor(hOutConsole, "red")
  
  SendText "                 Consoles Application from VB. (by pramod kumar) " & vbCrLf
  SendText "                 To get help type help and enter" & vbCrLf
  SendText "                 to exit type" & vbCrLf
  SendText "                 ""exit"" with out quotes" & vbCrLf
  Buffer = "                 " & String(50, "*")
  SendText Buffer
  Buffer = ""
  SendText vbCrLf & sPrompt
  
  Do
      Length = 0
      If Not WaitForSingleObject(hInConsole, 0) Then
          ReadConsoleInput hInConsole, myEvent, 1, Length
          Select Case myEvent.EventType
              Case CON_KEY_EVENT
                  If myEvent.KeyEvent.bKeyDown Then 'KeyDown
                      If myEvent.KeyEvent.uChar > 0 Then 'RealKey
                          If myEvent.KeyEvent.uChar = 13 Then
                              Quit = MyCommand(Buffer)
                              Buffer = ""
                          ElseIf myEvent.KeyEvent.uChar = 8 Then
                              If Buffer <> "" Then
                                  Buffer = left(Buffer, Len(Buffer) - 1)
                                  If right(Buffer, 1) = Chr(9) Then
                                      SendText String(80, Chr(8))
                                      SendText Buffer
                                  Else
                                      SendText Chr(8)
                                      SendText " "
                                      SendText Chr(8)
                                  End If
                              End If
                          Else
                              SendText Chr(myEvent.KeyEvent.uChar)
                              Buffer = Buffer & Chr(myEvent.KeyEvent.uChar)
                          End If
                      End If
                  End If
              Case CON_MOUSE_EVENT
              Case CON_WINDOW_BUFFER_SIZE_EVENT
              Case CON_MENU_EVENT
              Case CON_FOCUS_EVENT
          End Select
      End If
      
  Loop Until Quit
  
  CloseConsole
End Sub

Public Function OpenConsole(Capt As String)
  If AllocConsole() Then
      SetConsoleTitle Capt
      hInConsole = GetStdHandle(STD_INPUT_HANDLE)
      hOutConsole = GetStdHandle(STD_OUTPUT_HANDLE)
      If hOutConsole = 0 Then
          FreeConsole
      Else
          SetConsoleCtrlHandler AddressOf ConsoleHandler, True
          OpenConsole = True
      End If
  End If
End Function
Public Function ConsoleHandler(ByVal CtrlType As Long) As Long
  ConsoleHandler = 1  'This tells the console window to ignore all console
                      'signals. If you don't do this, closing the console window
                      'or typing Ctrl-Break would cause your program to end.
End Function
Public Function CloseConsole() As Boolean
  CloseHandle hInConsole
  CloseHandle hOutConsole
  FreeConsole
End Function
Public Function SendText(Text As String, Optional User As Boolean = True)
  Dim cWritten As Long
  If Not User Then
      Text = Text & vbCrLf & sPrompt
  End If
  
  SendText = WriteConsole(hOutConsole, ByVal Text, Len(Text), cWritten, ByVal 0&)
   
End Function
Public Function ErrDes(ErrNo As Long) As String
  Dim Buffer As String
  Buffer = Space(200)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, ErrNo, LANG_NEUTRAL, Buffer, 200, ByVal 0&
  ErrDes = Buffer
End Function
Public Function MyCommand(CmdLine As String) As Boolean
Dim Cmd As String, Arg As String
Dim L, i As Integer
Dim Rtn As Boolean
If CmdLine = "exit" Then
    MyCommand = True
Else
    L = Len(CmdLine)
    i = InStr(1, CmdLine, " ", vbTextCompare)
    If i Then
        Cmd = Mid(CmdLine, 1, i - 1)
        Arg = Mid(CmdLine, i + 1, L)
        Rtn = ProcessCommand(Cmd, Arg)
    Else
        Select Case CmdLine
            Case "cls"
                ClrScr hOutConsole
                Rtn = True
            Case "help"
                SendText vbCrLf & "cls "
                SendText vbCrLf & "clear the the console window"
                SendText vbCrLf & "dir path"
                SendText vbCrLf & "list the content of path and its sub folders"
                SendText vbCrLf & "set sysvariable newvalue"
                SendText vbCrLf & "systemvariable are"
                SendText vbCrLf & "bkcolor  - Background color of console"
                SendText vbCrLf & "txtcolor - Text color "
                SendText vbCrLf & "prompt -   Command prompt "
                SendText vbCrLf & "newvalue are"
                SendText vbCrLf & "red,green,blue"
                Rtn = True
            Case Else
                Rtn = False
        End Select
    End If
    If Not Rtn Then
        SendText vbCrLf & "Bad Command or file name", False
    Else
        SendText "", False
    End If
    MyCommand = False

End If
End Function

Public Function ProcessCommand(Cmd As String, Arg As String) As Boolean
Dim ArgCount As Integer, i As Integer, L As Integer, Start As Integer
Dim Rtn As Long
Dim Args() As String
L = Len(Arg)
ArgCount = 0
Start = 1
i = 1
For i = i To L
    Rtn = InStr(Start, Arg, " ", vbTextCompare)
    If Rtn Then
        ReDim Preserve Args(ArgCount) As String
        Args(ArgCount) = Mid(Arg, Start, Rtn - Start)
        ArgCount = ArgCount + 1
        Start = Rtn + 1
        i = Rtn
    Else

        ReDim Preserve Args(ArgCount) As String
        Args(ArgCount) = Mid(Arg, Start, L + 1 - Start)
        i = L
    End If
Next

Select Case Cmd
    
    Case "set"
        If ArgCount <> 1 Then
            SendText vbCrLf & "Too many parameter", False
        Else
            Select Case Args(0)
                Case "textcolor"
                    SetTextColor hOutConsole, Args(1)
                Case "bkcolor"
                    SetBkColor hOutConsole, Args(1)
                Case "prompt"
                    SetPrompt Args(1)
            End Select
        End If
        ProcessCommand = True
    Case "dir"
        If ArgCount <> 0 Then
            SendText vbCrLf & "Too many parameter", False
        Else
            L = Len(Args(0))
            Dim Rev As String, sFind As String, sDir As String
            Rev = StrReverse(Args(0))
            Rtn = InStr(1, Rev, "\", vbTextCompare)
            sFind = Mid(Rev, 1, Rtn - 1)
            sDir = Mid(Rev, Rtn, L - Rtn + 1)
            sDir = StrReverse(sDir)
            sFind = StrReverse(sFind)
            List sDir, sFind, ArgCount, Args
            ProcessCommand = True
        End If
    Case "else"
        ProcessCommand = False
End Select
End Function

