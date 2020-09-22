VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "End Selected Process"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "List Running Processes"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Process ID Number"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Process Name:  Process ID#"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'You can get CODE FIXER at: 53297&lngWId=1

Option Explicit

Private Type PROCESSENTRY32
    dwSize                  As Long
    cntUsage                As Long
    th32ProcessID           As Long
    th32DefaultHeapID       As Long
    th32ModuleID            As Long
    cntThreads              As Long
    th32ParentProcessID     As Long
    pcPriClassBase          As Long
    dwFlags                 As Long
    szExeFile               As String * 260
End Type

Private Const ParseMe   As String = ""    'I'll use this variable (chr(1)) to split up a string.

Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32.DLL" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "KERNEL32.DLL" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "KERNEL32.DLL" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "KERNEL32.DLL" (ByVal hHandle As Long) As Long
Private Declare Function OpenProcess Lib "KERNEL32.DLL" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "KERNEL32.DLL" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Function CleanItUp(ByVal strOrig As String) As String

    CleanItUp = Left$(strOrig, InStr(strOrig, vbNullChar) - 1)
     'shortens the original by stopping at the first vbNullChar it comes to

End Function

Private Sub Command1_Click()

Dim pProcesses As String   'declare variables
Dim lList()    As String
Dim I          As Long
    List1.Clear 'Clear List1 of any existing information
    pProcesses = GetTheProcesses 'see the function GetTheProcesses
    lList = Split(pProcesses, ParseMe) 'adds data to the string array as needed
    For I = 0 To (UBound(lList) - 1) 'starts a loop so we can add to List1
        List1.AddItem lList(I) & ":  " & lList(I + 1) 'adds to List1 what's held in the array
        I = I + 1
         'We've already used I + 1 in the above line.
             'There's a better way of doing this with a ListView control, but that's not what
             'this Demo is for.  I wanted to keep it as simple as possible.
    Next I 'next loop

End Sub

Private Sub Command2_Click()

Dim PID As Long 'declare variables

    PID = Label2.Caption 'set Variable
    KillProcess PID 'Passes the Process ID to the function with the variable

End Sub

Private Function GetTheProcesses() As String

Dim pProcess  As PROCESSENTRY32  'declare variables
Dim sSnapShot As Long
Dim rReturn   As Integer

    sSnapShot = CreateToolhelp32Snapshot(15, 0) 'setting variables
    pProcess.dwSize = Len(pProcess)
    Process32First sSnapShot, pProcess 'gets the first process ([System Process])
    Do 'starts another loop
        GetTheProcesses = GetTheProcesses & CleanItUp(pProcess.szExeFile) & ParseMe & pProcess.th32ProcessID & ParseMe 'adds to the string variable the next process and ID
        rReturn = Process32Next(sSnapShot, pProcess)
         'gets the next process so we know to loop again
        DoEvents 'free's up the computer
    Loop While rReturn <> 0 'as long as there's a next process, we'll loop again
    CloseHandle sSnapShot 'frees up the handle

End Function

Private Sub KillProcess(pProcessID As Long)

    TerminateProcess OpenProcess(2035711, 1, pProcessID), 0 'Ends the process selected
    DoEvents 'Let's the computer "do it's stuff"
    Command1 = True 'basically this part just updates List1 by calling a Command1.Click = true

End Sub

Private Sub List1_Click()

Dim sSeperator As Long   'declare variables
Dim sSelNum    As Long
Dim sSelected  As String

    sSelNum = List1.ListIndex 'set variables
    sSelected = List1.List(sSelNum)
    sSeperator = InStr(1, sSelected, ":") 'looks for the : in the selected item from List1
    Label2.Caption = Mid$(sSelected, sSeperator + 3, Len(sSelected))
     'Sets the caption of the label to only the Process ID selected
    Command2.Enabled = Not (LenB(Label2.Caption) = 0 Or Label2.Caption = "0") 'If Label2.Caption
     '<> nothing or zero, it's enabled.

End Sub

