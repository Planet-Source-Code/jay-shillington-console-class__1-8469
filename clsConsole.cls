VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'based upon the DOS Console posted on http://www.planet-source-code.com/  by Loreno Heer

Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleOutput As Long, dwMode As Long) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const STD_INPUT_HANDLE = -10&
Private Const STD_OUTPUT_HANDLE = -11&
Private Const STD_ERROR_HANDLE = -12&
Private Const FOREGROUND_RED = &H4
Private Const FOREGROUND_GREEN = &H2
Private Const FOREGROUND_BLUE = &H1
Private Const FOREGROUND_INTENSITY = &H8
Private Const BACKGROUND_RED = &H40
Private Const BACKGROUND_GREEN = &H20
Private Const BACKGROUND_BLUE = &H10
Private Const BACKGROUND_INTENSITY = &H80
Private Const ENABLE_LINE_INPUT = &H2
Private Const ENABLE_ECHO_INPUT = &H4
Private Const ENABLE_MOUSE_INPUT = &H10
Private Const ENABLE_PROCESSED_INPUT = &H1
Private Const ENABLE_WINDOW_INPUT = &H8
Private Const ENABLE_PROCESSED_OUTPUT = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT = &H2

Private hConsoleIn As Long
Private hConsoleOut As Long
Private hConsoleErr As Long

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_SHOWWINDOW = &H40

Private Const WM_CLOSE = &H10
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_RESTORE = 9

'local variable(s) to hold property value(s)
Private mvarLogFilePathName As String 'local copy
Private mvarConsoleWindowTitle As String 'local copy

Public Property Let ConsoleWindowTitle(ByVal vData As String)

    mvarConsoleWindowTitle = vData

End Property


Public Property Get ConsoleWindowTitle() As String

    ConsoleWindowTitle = mvarConsoleWindowTitle

End Property


Public Sub WriteOut(ByVal Msg As String, Optional ByVal LogIt As Boolean = False)

  Msg = Format(Now, "hh:nn:ss") & " " & Msg
  
  Msg = Msg & vbCrLf
  
  WriteConsole hConsoleOut, Msg, Len(Msg), vbNull, vbNull
  
  If LogIt Then WriteLog Msg
    
End Sub


Public Sub Important(ByVal Msg As String, Optional ByVal LogIt As Boolean = False)
  
  SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_INTENSITY
  
  WriteOut "---------------------------------------------------------------------", LogIt
  WriteOut Msg, LogIt
  WriteOut "---------------------------------------------------------------------", LogIt
  
  SetConsoleTextAttribute hConsoleOut, FOREGROUND_RED Or FOREGROUND_GREEN Or FOREGROUND_BLUE
  
End Sub


Public Property Let LogFilePathName(ByVal vData As String)
    
  mvarLogFilePathName = vData

End Property


Public Property Get LogFilePathName() As String
    
  LogFilePathName = mvarLogFilePathName

End Property


Private Function CGet() As String
   
  Dim sUserInput As String * 256
  Call ReadConsole(hConsoleIn, sUserInput, Len(sUserInput), vbNull, vbNull)
  CGet = Left$(sUserInput, InStr(sUserInput, Chr$(0)) - 3)

End Function


Public Sub CloseConsole()
   
  FreeConsole
   
End Sub


Public Sub LoadConsole()
   
  Dim lHwnd As Long
  Dim ConsoleTitle As String
  
  ConsoleTitle = ConsoleWindowTitle()
  
  AllocConsole
  SetConsoleTitle ConsoleTitle
  hConsoleIn = GetStdHandle(STD_INPUT_HANDLE)
  hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE)
  hConsoleErr = GetStdHandle(STD_ERROR_HANDLE)

  lHwnd = FindWindow("ConsoleWindowClass", ConsoleTitle)
  
  ShowWindow lHwnd, SW_SHOWNORMAL
  Call SetWindowPos(lHwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  Call SetWindowPos(lHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    
End Sub


Private Sub WriteLog(sMsg As String)

  If LogFilePathName = "" Then Exit Sub
  
  Dim intFile As Integer   ' FreeFile variable
  
  intFile = FreeFile()
  
  Open LogFilePathName For Append As #intFile
  Print #intFile, sMsg
  Close #intFile
  
End Sub


Private Function FileExists(FullPathandFile As String) As Boolean

  On Error Resume Next
  
  If FileLen(FullPathandFile) > 0& Then
    If Err = 0 Then FileExists = True
  End If

End Function


Private Sub Class_Terminate()

  CloseConsole
  
End Sub
