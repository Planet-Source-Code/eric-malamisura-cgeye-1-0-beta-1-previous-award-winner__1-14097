VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTemplate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Templates / Wizards"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3825
      TabIndex        =   1
      Top             =   3465
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   0
      TabIndex        =   3
      Top             =   15
      Width           =   4980
      _ExtentX        =   8784
      _ExtentY        =   5953
      Arrange         =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.FileListBox filePlugin 
      Height          =   285
      Left            =   5475
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   2670
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "template.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "template.frx":0F84
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label pluginname 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   4725
   End
End
Attribute VB_Name = "frmTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

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
      Dim addtemplate
 Dim i As Integer
 Screen.MousePointer = 11
  pathname = App.Path & "\Templates\"
MyPath = pathname
MYNAME = Dir(MyPath & "*.cgi", vbNormal)
Do While Not MYNAME = ""
    If Not MYNAME = "." And Not MYNAME = ".." Then
If (GetAttr(MyPath & MYNAME) And vbNormal) = vbNormal Then
  
      Set addtemplate = ListView1.ListItems.Add(1, , MYNAME, 2, 2)
        End If

    End If
    MYNAME = Dir


Loop

ListView1.Refresh
Screen.MousePointer = 0

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

filePlugin.Path = App.Path & "\Wizards\"
filePlugin.Pattern = "*.dll"

For a = 0 To filePlugin.ListCount - 1
    strTemp = Left(filePlugin.List(a), InStr(filePlugin.List(a), ".") - 1)
    Set MyObj = CreateObject(strTemp & ".Main")
      Dim addwizard
      Set addwizard = ListView1.ListItems.Add(1, , ExtractFileName("" & filePlugin.List(a)), 1, 1)
    Set MyObj = Nothing
Next
Exit Sub
handler:
If Err.Number = 429 Then
    RunShell ("regsvr32 /s " & CorrectPath(filePlugin.Path) & filePlugin.List(a))
    Err.Clear
    Resume
Else
    MsgBox Err.Number & " - " & Err.Description, vbCritical
End If

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

Private Sub lblInfo_Click(Index As Integer)

End Sub

Private Sub ListView1_DblClick()
On Error GoTo handler:
Me.Hide
strTemp = App.Path & "\Templates\" & ListView1.SelectedItem
Set MyObj = CreateObject(ListView1.SelectedItem & ".Main")

Do While MyObj.WizardStart = 1
    DoEvents
Loop

If MyObj.strReturn <> "" Then
mdiMain.NewDocument
mdiMain.ActiveForm.txtMain.SelText = MyObj.strReturn
End If
If MyObj.strfilename <> "" Then
mdiMain.ActiveForm.Text1.Text = MyObj.strfilename
mdiMain.ActiveForm.Caption = MyObj.strfilename
End If
Set MyObj = Nothing
Exit Sub

handler:
If Err.Number > "429" Then
MsgBox Err.Description
Err.Clear
Else
Err.Clear
On Error GoTo Error:
mdiMain.NewDocument
mdiMain.ActiveForm.txtMain.LoadFile strTemp

End If
Error:
If Err Then
'MsgBox Err.Description
End If
End Sub

Public Function ExtractFileName(strPath As String) As String
    If strPath = "*.frm" Then
    MsgBox "Insert Template"
    Else
    strPath = strPath
    ExtractFileName = Left(strPath, InStr(strPath, ".") - 1)
    'ExtractFileName = StrReverse(strPath)
    End If
End Function


