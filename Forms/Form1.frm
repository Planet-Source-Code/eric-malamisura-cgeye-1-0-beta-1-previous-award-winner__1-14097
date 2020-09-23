VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4875
      TabIndex        =   5
      Top             =   3195
      Width           =   1095
   End
   Begin VB.ListBox lstPlugins 
      Height          =   1620
      Left            =   120
      TabIndex        =   2
      Top             =   795
      Width           =   5865
   End
   Begin VB.FileListBox filePlugin 
      Height          =   285
      Left            =   5475
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label pluginname 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   4725
   End
   Begin VB.Label lblDescr 
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   75
      Width           =   3720
   End
   Begin VB.Label lblInfo 
      Caption         =   "Plugin Info :"
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblPluginInfo 
      BackStyle       =   0  'Transparent
      Height          =   585
      Left            =   105
      TabIndex        =   3
      Top             =   2985
      Width           =   4695
   End
   Begin VB.Label lblInfo 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   495
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const PROCESS_QUERY_INFORMATION = &H400
Const STATUS_PENDING = &H103&  'new
''''''''

Public MyPath As String

Dim a As Long 'Temp counter..
Dim strTemp As String 'temp string
Dim MyObj As Object 'temp object for loading the plugins...

Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub Form_Load()

MyPath = App.Path

lblDescr = "Single click on a plugin to view its version, info and name, double click to run it."
lblInfo(0) = "reading available plugins"

Call GetAvailPlugins

End Sub

Private Function CorrectPath(strPath As String) As String

If Len(strPath) = 3 Then
    CorrectPath = strPath
Else
    CorrectPath = strPath & "\"
End If

End Function

Private Sub GetAvailPlugins()


On Error GoTo handler:

filePlugin.Path = App.Path & "\Wizards"
filePlugin.Pattern = "*.dll"

For a = 0 To filePlugin.ListCount - 1
    'just get the name of the file without the .dll extension..
    lblInfo(0) = "Reading " & CorrectPath(filePlugin.Path) & filePlugin.List(a)
    strTemp = Left(filePlugin.List(a), InStr(filePlugin.List(a), ".") - 1)
    Set MyObj = CreateObject(strTemp & ".Main")
    'add the plugin to the menu..
    Call AddToPluginMenu(MyObj.pluginId, strTemp & ".Main")
    strTemp = MyObj.pluginId & " - " & MyObj.pluginversion & " - " & MyObj.plugindescription
    lstPlugins.AddItem filePlugin.List(a)
    Set MyObj = Nothing
Next

lblInfo(0).Visible = False

Exit Sub

handler:

If Err.Number = 429 Then 'ActiveX component can't create object..
    'we will attempt to register the dll..
    RunShell ("regsvr32 /s " & CorrectPath(filePlugin.Path) & filePlugin.List(a))
    Err.Clear
    Resume
Else
    MsgBox Err.Number & " - " & Err.Description, vbCritical
End If



End Sub

Private Sub AddToPluginMenu(strPlugin As String, strKey As String)

End Sub

Private Sub RunShell(cmdline$)
Dim hProcess As Long
Dim ProcessId As Long
Dim exitCode As Long
Dim r As Long

    ProcessId& = Shell(cmdline$, 1)
    hProcess& = OpenProcess(PROCESS_QUERY_INFORMATION, False, ProcessId&)
    Do
        Call GetExitCodeProcess(hProcess&, exitCode&)
        DoEvents
    Loop While exitCode& = STATUS_PENDING
    If exitCode& <> 0 Then
        MsgBox "Error Registering: " & cmdline$, vbExclamation, "ERROR"
    End If
    r = CloseHandle(hProcess)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'clean up the users system by un-registering the dlls..
'I'm nice ;-)

End Sub

Private Sub lstPlugins_Click()

'get the plugin info...PluginName
strTemp = Left(lstPlugins.List(lstPlugins.ListIndex), InStr(lstPlugins.List(lstPlugins.ListIndex), ".") - 1)
Set MyObj = CreateObject(strTemp & ".Main")

strTemp = "Plugin ID       : " & MyObj.pluginId & vbCrLf & _
    "Version         : " & MyObj.pluginversion & vbCrLf & _
    "Description   : " & MyObj.plugindescription

lblPluginInfo = strTemp

Set MyObj = Nothing

End Sub

Private Sub lstPlugins_DblClick()

On Error GoTo handler:
Me.Hide
'run the plugins PluginStart Function...
strTemp = Left(lstPlugins.List(lstPlugins.ListIndex), InStr(lstPlugins.List(lstPlugins.ListIndex), ".") - 1)
Set MyObj = CreateObject(strTemp & ".Main")

Do While MyObj.pluginstart = 1
    DoEvents
Loop

'get the information back from the plugin dll..
If MyObj.strReturn <> "" Then
    'MsgBox "Returned from plugin : " & MyObj.strReturn
    mdiMain.NewDocument
    mdiMain.ActiveForm.InsertTag MyObj.strReturn, ""
End If

Set MyObj = Nothing

Exit Sub

handler:
MsgBox Err.Number & " - " & Err.Description, vbExclamation

End Sub

