Attribute VB_Name = "mdlSet"
'Filename: mdlSet.bas
'Desc    : This module contain functions set command
'Author  : pramod kumar
'E-mail  :tk_pramod@yahoo.com
'Company : Imprudents (www.imprudents.com)
'Copyright © 2000-2001 Imprudents, All Rights Reserved
'Date    :11-Feb-2001

Public Sub SetPrompt(Promp As String)
sPrompt = Promp
End Sub

Public Sub ClrScr(hconsole As Long)

  Dim coordScreen As COORD
  coordScreen.x = 0
  coordScreen.y = 0  '/* here's where we'll home the  cursor */
  Dim bSuccess As Long
  Dim cCharsWritten As Long
  Dim csbi As CONSOLE_SCREEN_BUFFER_INFO ' /* to get buffer info */
  Dim dwConSize As Long                  '/* number of character cells in
  Dim cc As CONSOLE_CURSOR_INFO          '    the current buffer */
  Dim lngCoord As Long
  Dim lngWritten As Long
  
  '/* get the number of character cells in the current buffer */

  bSuccess = GetConsoleScreenBufferInfo(hconsole, csbi)
  
  dwConSize = csbi.dwSize.x * csbi.dwSize.y
  lngCoord = cvtCoordToLng(coordScreen.y, coordScreen.x)
  '/* fill the entire screen with blanks */
  Dim x As Integer
  
   bSuccess = FillConsoleOutputCharacter(hconsole, 32&, dwConSize, lngCoord, lngWritten)
  '/* get the current text attribute */
  bSuccess = GetConsoleScreenBufferInfo(hconsole, csbi)
  '/* now set the buffer's attributes accordingly */
  bSuccess = FillConsoleOutputAttribute(hconsole, csbi.wAttributes, dwConSize, lngCoord, cCharsWritten)
  '/* put the cursor at (0, 0) */
  coordScreen.x = 0
  coordScreen.y = 0
  lngCoord = cvtCoordToLng(coordScreen.y, coordScreen.x)
  bSuccess = SetConsoleCursorPosition(hconsole, lngCoord)
End Sub
Private Function cvtCoordToLng(wHi As Integer, wLo As Integer) As Long
  cvtCoordToLng = (wHi * &H10000) Or (wLo And &HFFFF&)
End Function
Public Function SetBkColor(hconsole As Long, Color As String) As Boolean
Dim cWritten As Long
Dim fSuccess As Long
Dim COORD As COORD
Dim lngColor As Long
COORD.x = 0            '// start at first cell
COORD.y = 0            '//   of first row
Dim csbi As CONSOLE_SCREEN_BUFFER_INFO ' /* to get buffer info */
fSuccess = GetConsoleScreenBufferInfo(hconsole, csbi)
Select Case Color
  Case "green"
      lngColor = BACKGROUND_GREEN Or FOREGROUND_GREEN Or FOREGROUND_BLUE Or FOREGROUND_RED
  Case "red"
      lngColor = BACKGROUND_RED Or FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE
  Case "blue"
      lngColor = BACKGROUND_BLUE Or FOREGROUND_BLUE Or FOREGROUND_GREEN Or FOREGROUND_RED
  Case Else
     
End Select
  
fSuccess = FillConsoleOutputAttribute(hconsole, lngColor, 80 * 50, cvtCoordToLng(COORD.x, COORD.y), cWritten)
End Function
Public Function SetTextColor(hOutConsole As Long, Color As String) As Boolean
Dim lngColor As Long
Select Case Color
  Case "green"
      lngColor = FOREGROUND_GREEN
  Case "red"
      lngColor = FOREGROUND_RED
  Case "blue"
      lngColor = FOREGROUND_BLUE
  Case Else
      
End Select

SetConsoleTextAttribute hOutConsole, lngColor
SetTextColor = True
End Function
