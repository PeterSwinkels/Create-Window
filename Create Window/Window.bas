Attribute VB_Name = "WindowModule"
'This module contains this program's core procedures.
Option Explicit

'The Micrsoft Windows API constants, functions, and structures used by this program.
Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type MSG
   hwnd As Long
   Message As Long
   wParam As Long
   lParam As Long
   time As Long
   pt As POINTAPI
End Type

Private Type WNDCLASSEX
   cbSize As Long
   Style As Long
   lpfnWndProc As Long
   cbClsExtra As Long
   cbWndExtra As Long
   hInstance As Long
   hIcon As Long
   hCursor As Long
   hbrBackground As Long
   lpszMenuName As String
   lpszClassName As String
   hIconSm As Long
End Type

Private Const COLOR_BTNFACE As Long = 15&
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_SUCCESS As Long = 0&
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const SW_SHOWNORMAL As Long = 1&
Private Const WM_CLOSE As Long = &H10&
Private Const WM_COMMAND As Long = &H111&
Private Const WM_DESTROY As Long = &H2&
Private Const WM_GETTEXT As Long = &HD&
Private Const WM_GETTEXTLENGTH As Long = &HE&
Private Const WM_PAINT As Long = &HF&
Private Const WM_SETTEXT As Long = &HC&
Private Const WS_BORDER As Long = &H800000
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_CONTROLPARENT As Long = &H10000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_TABSTOP As Long = &H10000
Private Const WS_VISIBLE As Long = &H10000000

Private Declare Function CreateWindowExA Lib "User32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Long) As Long
Private Declare Function DefWindowProcA Lib "User32.dll" (ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function DispatchMessageA Lib "User32.dll" (lpMsg As MSG) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetDC Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetMessageA Lib "User32.dll" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function IsDialogMessageA Lib "User32.dll" (ByVal hDlg As Long, lpMsg As MSG) As Long
Private Declare Function IsWindow Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function RegisterClassExA Lib "User32.dll" (lpwcx As WNDCLASSEX) As Long
Private Declare Function ReleaseDC Lib "User32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function SendMessageA Lib "User32.dll" (ByVal hwnd As Long, ByVal MSG As Long, wParam As Long, lParam As Long) As Long
Private Declare Function SetFocus Lib "User32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "User32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function TextOutA Lib "Gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function TranslateMessage Lib "User32.dll" (lpMsg As MSG) As Long
Private Declare Function UnregisterClassA Lib "User32.dll" (ByVal lpClassName As String, ByVal hInstance As Long) As Long

'The constants and variables used by this program.
Private Const GET_MESSAGE_ERROR As Long = -1   'Defines an error while calling the GetMessageA API function.
Private Const GET_MESSAGE_QUIT As Long = 0     'Indicates that the window has been closed while calling the GetMessageA API function.
Private Const NO_HANDLE As Long = 0            'Defines a null handle.
Private Const MAX_STRING As Long = 65535       'Defines the maximum number of characters used for a string buffer.

Private ChangeButtonH As Long   'Contains the "Change" button handle.
Private MainWindowH As Long     'Contains the main window handle.
Private QuitButtonH As Long     'Contains the "Quit" button handle.
Private TextBoxH As Long        'Contains the text box's handle.

'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long) As Long
On Error GoTo ErrorTrap
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long
Dim Message As String

   ErrorCode = Err.LastDllError
   Err.Clear
   
   If Not (ErrorCode = ERROR_SUCCESS Or ErrorCode = ERROR_ACCESS_DENIED) Then
      Description = String$(MAX_STRING, vbNullChar)
      Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
      If Length = 0 Then
         Description = "No description."
      ElseIf Length > 0 Then
         Description = Left$(Description, Length - 1)
      End If
     
      Message = "API error code: " & CStr(ErrorCode) & " - " & Description
      Message = Message & "Return value: " & CStr(ReturnValue) & vbCr
      If MsgBox(Message, vbExclamation Or vbOKCancel) = vbCancel Then End
   End If
   
EndRoutine:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure creates the window and sets its properties.
Private Sub CreateProgramWindow()
On Error GoTo ErrorTrap
Dim WindowClass As WNDCLASSEX
Dim x As Long
Dim y As Long

   With WindowClass
      .cbSize = Len(WindowClass)
      .hbrBackground = (COLOR_BTNFACE + 1)
      .hInstance = App.hInstance
      .lpfnWndProc = GetAddress(AddressOf EventHandler)
      .lpszClassName = "STATIC"
   End With
   
   If CBool(CheckForError(RegisterClassExA(WindowClass))) Then
      x = CLng(((Screen.Width / Screen.TwipsPerPixelX) / 2) - 188)
      y = CLng(((Screen.Height / Screen.TwipsPerPixelY) / 3) - 80)
      
      MainWindowH = CheckForError(CreateWindowExA(WS_EX_CONTROLPARENT, "STATIC", "Window:", WS_SYSMENU Or WS_VISIBLE, x, y, 374, 160, CLng(0), CLng(0), App.hInstance, CLng(0)))
      If Not MainWindowH = NO_HANDLE Then
         TextBoxH = CheckForError(CreateWindowExA(CLng(0), "EDIT", vbNullString, WS_BORDER Or WS_CHILD Or WS_TABSTOP Or WS_VISIBLE, 24, 40, 320, 24, MainWindowH, CLng(0), App.hInstance, CLng(0)))
         ChangeButtonH = CheckForError(CreateWindowExA(CLng(0), "BUTTON", "&Change", WS_CHILD Or WS_TABSTOP Or WS_VISIBLE, 24, 76, 128, 32, MainWindowH, CLng(0), App.hInstance, CLng(0)))
         QuitButtonH = CheckForError(CreateWindowExA(CLng(0), "BUTTON", "&Quit", WS_CHILD Or WS_TABSTOP Or WS_VISIBLE, 216, 76, 128, 32, MainWindowH, CLng(0), App.hInstance, CLng(0)))
         
         CheckForError ShowWindow(MainWindowH, SW_SHOWNORMAL)
         CheckForError SetFocus(TextBoxH)
      End If
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This function displays the specified text in the specified window.
Private Sub DisplayText(WindowH As Long, x As Long, y As Long, Text As String)
On Error GoTo ErrorTrap
Dim WindowDC As Long

   WindowDC = CheckForError(GetDC(WindowH))
   CheckForError TextOutA(WindowDC, x, y, Text, Len(Text))
   CheckForError ReleaseDC(WindowH, WindowDC)
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure processes the window's events and gives the command to do default event handling.
Private Function EventHandler(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo ErrorTrap
Dim ReturnValue As Long

   ReturnValue = CheckForError(DefWindowProcA(hwnd, uMsg, wParam, lParam))
   
   Select Case uMsg
      Case WM_COMMAND
         Select Case lParam
            Case ChangeButtonH
               CheckForError SendMessageA(MainWindowH, WM_SETTEXT, CLng(0), ByVal StrPtr(GetWindowText(TextBoxH)))
            Case QuitButtonH
               CheckForError SendMessageA(hwnd, WM_CLOSE, CLng(0), CLng(0))
         End Select
      Case WM_CLOSE
         CheckForError UnregisterClassA("STATIC", App.hInstance)
      Case WM_DESTROY
         If CBool(CheckForError(IsWindow(ChangeButtonH))) Then CheckForError DestroyWindow(ChangeButtonH)
         If CBool(CheckForError(IsWindow(QuitButtonH))) Then CheckForError DestroyWindow(QuitButtonH)
         If CBool(CheckForError(IsWindow(TextBoxH))) Then CheckForError DestroyWindow(TextBoxH)
         If CBool(CheckForError(IsWindow(MainWindowH))) Then CheckForError DestroyWindow(MainWindowH)
      Case WM_PAINT
         DisplayText MainWindowH, 24, 16, "Enter new window title:"
   End Select
EndRoutine:
   
   EventHandler = ReturnValue
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the address proceduced by the AddressOf operator.
Private Function GetAddress(Address As Long) As Long
On Error GoTo ErrorTrap
EndRoutine:
   GetAddress = Address
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the specified window's text.
Private Function GetWindowText(WindowH As Long) As String
On Error GoTo ErrorTrap
Dim Length As Long
Dim Text As String
   
   Length = CheckForError(SendMessageA(WindowH, WM_GETTEXTLENGTH, CLng(0), CLng(0))) + 1
   Text = String$(Length, vbNullChar)
   CheckForError SendMessageA(WindowH, WM_GETTEXT, ByVal Length, ByVal StrPtr(Text))
   
EndRoutine:
   GetWindowText = Text
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Private Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   
   On Error Resume Next
   If MsgBox(Description & vbCr & "Error code: " & ErrorCode, vbExclamation Or vbOKCancel) = vbCancel Then End
End Sub


'This procedure instructs the program to create a window.
Private Sub Main()
On Error GoTo ErrorTrap
Dim Message As MSG
Dim ReturnValue As Long

   CreateProgramWindow
   
   Do While CBool(IsWindow(MainWindowH))
      ReturnValue = CheckForError(GetMessageA(Message, MainWindowH, CLng(0), CLng(0)))
   
      If ReturnValue = GET_MESSAGE_ERROR Or ReturnValue = GET_MESSAGE_QUIT Then
         Exit Do
      ElseIf Not CBool(CheckForError(IsDialogMessageA(MainWindowH, Message))) Then
         CheckForError TranslateMessage(Message)
         CheckForError DispatchMessageA(Message)
      End If
   Loop
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

