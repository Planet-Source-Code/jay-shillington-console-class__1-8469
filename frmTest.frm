VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Sample Console Application"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   8400
   Begin VB.CheckBox ckLogIt 
      Caption         =   "&Use Log File"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1860
      Width           =   2535
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "E&xit"
      Default         =   -1  'True
      Height          =   315
      Index           =   2
      Left            =   6840
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtLogFile 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5415
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Important, Write to console"
      Height          =   315
      Index           =   1
      Left            =   4980
      TabIndex        =   3
      Top             =   720
      Width           =   3315
   End
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   8235
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Write to console"
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   720
      Width           =   3315
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Log file path and name:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   1665
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Text to write to console/log file:"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2220
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Console As New clsConsole

Private Sub cmdBtn_Click(Index As Integer)

  Const btnNormal = 0
  Const btnImportant = 1
  Const btnExit = 2
  
  Select Case Index
  Case btnNormal
    Console.WriteOut txt, GetLogState()
  Case btnImportant
    Console.Important txt, GetLogState()
  Case btnExit
    Set Console = Nothing
    End
  End Select
  
End Sub

Private Sub Form_Load()
  
  Me.Top = 0
  Me.Left = 0
  
  Console.ConsoleWindowTitle = "Console Test"
  
  Console.LoadConsole
  
  ShowWelcome

  
End Sub

Private Sub txtLogFile_Change()

  If Len(txtLogFile) = 0 Then
    ckLogIt.Enabled = False
    ckLogIt.Value = vbUnchecked
  Else
    ckLogIt.Enabled = True
  End If
  
End Sub

Function GetLogState() As Boolean

  Dim bRet As Boolean
  
  bRet = False
  
  If ckLogIt.Value = vbChecked Then
    bRet = True
    Console.LogFilePathName = txtLogFile
  End If
  
  GetLogState = bRet
      
End Function

Sub ShowWelcome()

' created using http://st-www.cs.uiuc.edu/users/chai/figlet.html

Console.WriteOut String(79, "=")
Console.WriteOut "            __          __ "
Console.WriteOut "            \ \        / /     | |"
Console.WriteOut "             \ \  /\  / / ___  | |   ___   ___    _ __ ___     ___"
Console.WriteOut "              \ \/  \/ / / _ \ | |  / __| / _ \  | '_ ` _ \   / _ \"
Console.WriteOut "               \  /\  / |  __/ | | | (__ | (_) | | | | | | | |  __/"
Console.WriteOut "                \/  \/   \___| |_|_ \___| \___/  |_| |_| |_|  \___|"
Console.WriteOut "                                 | |"
Console.WriteOut "                                 | |_   ___"
Console.WriteOut "                                 | __| / _ \"
Console.WriteOut "                 _____           | |_ | (_) | "
Console.WriteOut "                / ____|           \__| \___/         | |"
Console.WriteOut "               | |       ___    _ __    ___    ___   | |   ___"
Console.WriteOut "               | |      / _ \  | '_ \  / __|  / _ \  | |  / _ \"
Console.WriteOut "               | |____ | (_) | | | | | \__ \ | (_) | | | |  __/"
Console.WriteOut "                \_____| \___/  |_| |_| |___/  \___/  |_|  \___|"
Console.WriteOut "                       _____"
Console.WriteOut "                      |  __ \"
Console.WriteOut "                      | |  | |   ___   _ __ ___     ___"
Console.WriteOut "                      | |  | |  / _ \ | '_ ` _ \   / _ \"
Console.WriteOut "                      | |__| | |  __/ | | | | | | | (_) |"
Console.WriteOut "                      |_____/   \___| |_| |_| |_|  \___/"
Console.WriteOut " "
Console.WriteOut String(79, "=")
End Sub
