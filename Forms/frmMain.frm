VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "New Document"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8655
   FillColor       =   &H00C0C0C0&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   -15
      Top             =   6495
   End
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   6405
      Left            =   330
      TabIndex        =   0
      Top             =   285
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   11298
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   50000
      TextRTF         =   $"frmMain.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picLines 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      ForeColor       =   &H00FFFFFF&
      Height          =   6405
      Left            =   -15
      ScaleHeight     =   6345
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   285
      Width           =   360
   End
   Begin VB.PictureBox topbar 
      Align           =   1  'Align Top
      BackColor       =   &H80000004&
      FillColor       =   &H00808080&
      ForeColor       =   &H00808080&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   8595
      TabIndex        =   3
      Top             =   0
      Width           =   8655
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Left            =   5580
         TabIndex        =   8
         Top             =   30
         Width           =   2985
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   225
         Left            =   45
         TabIndex        =   5
         Text            =   "New Document"
         Top             =   30
         Width           =   6330
      End
      Begin VB.CommandButton topbar1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -30
         TabIndex        =   4
         Top             =   0
         Width           =   8640
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   165
         Left            =   7290
         TabIndex        =   7
         Top             =   45
         Width           =   540
      End
   End
   Begin RichTextLib.RichTextBox txtParameters 
      Height          =   945
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Smart Tips/Debug Window"
      Top             =   6735
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1667
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   2
      MousePointer    =   1
      RightMargin     =   5
      TextRTF         =   $"frmMain.frx":03EE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.FileListBox filePlugin 
      Height          =   870
      Left            =   5700
      TabIndex        =   6
      Top             =   945
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_newtemplate 
         Caption         =   "&New From Template"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnu_revert 
         Caption         =   "&Revert "
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_documents 
         Caption         =   "<No Documents>"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_redo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu_delete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_find 
         Caption         =   "&Find"
      End
      Begin VB.Menu mnu_gotoline 
         Caption         =   "&Goto Line"
      End
   End
   Begin VB.Menu mnu_insertmenu 
      Caption         =   "&Insert"
      Begin VB.Menu mnu_Forms 
         Caption         =   "&Forms"
         Begin VB.Menu mnu_insform 
            Caption         =   "For&m"
            Index           =   0
         End
         Begin VB.Menu mnu_insform 
            Caption         =   "&Button"
            Index           =   1
         End
         Begin VB.Menu mnu_insform 
            Caption         =   "&1 Line Text Box"
            Index           =   2
         End
         Begin VB.Menu mnu_insform 
            Caption         =   "&Multiline Text Box"
            Index           =   3
         End
         Begin VB.Menu mnu_insform 
            Caption         =   "&Check box"
            Index           =   4
         End
         Begin VB.Menu mnu_insform 
            Caption         =   "&Option button"
            Index           =   5
         End
         Begin VB.Menu mnu_insform 
            Caption         =   "&ComboBox"
            Index           =   6
         End
         Begin VB.Menu mnu_insform 
            Caption         =   "&Label"
            Index           =   7
         End
      End
      Begin VB.Menu mnu_tables 
         Caption         =   "&Tables"
      End
      Begin VB.Menu mnu_inserturl 
         Caption         =   "&URL"
      End
      Begin VB.Menu mnu_insertpicture 
         Caption         =   "&Picture"
      End
      Begin VB.Menu mnu_insertfont 
         Caption         =   "&Font"
      End
   End
   Begin VB.Menu mnu_script 
      Caption         =   "&Script"
      Begin VB.Menu mnu_CommentBlock 
         Caption         =   "&Comment Block"
      End
      Begin VB.Menu mnu_uncommentblock 
         Caption         =   "&Uncomment Block"
      End
      Begin VB.Menu mnu_line2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnu_wizard 
         Caption         =   "Wizards"
         Begin VB.Menu mnu_ifthen 
            Caption         =   "If .... Else"
         End
         Begin VB.Menu mnu_shebangline 
            Caption         =   "&Shebang Line"
         End
      End
      Begin VB.Menu mnu_line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Text 
         Caption         =   "&Text"
         Begin VB.Menu mnu_Bold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu mnu_Italic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu mnu_Underline 
            Caption         =   "&Underline"
         End
      End
      Begin VB.Menu mnu_insert 
         Caption         =   "&Insert Tags"
         Begin VB.Menu mnu_LeftTag 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnu_CenterTag 
            Caption         =   "&Center"
         End
         Begin VB.Menu mnu_RightTag 
            Caption         =   "&Right"
         End
         Begin VB.Menu line1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_printline 
            Caption         =   "&Print"
         End
      End
   End
   Begin VB.Menu mnuplugins 
      Caption         =   "&Plugins"
      Begin VB.Menu mnupluginoptions 
         Caption         =   "No Plugins Installed"
         Enabled         =   0   'False
         Index           =   0
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Begin VB.Menu mnu_toolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_formstoolbar 
         Caption         =   "&Insert Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_statusbar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_linenumbers 
         Caption         =   "&Line Numbers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_leftwin 
         Caption         =   "&Left Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu5_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_settings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_tilehorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnu_tilevertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnu_cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_arrangeicons 
         Caption         =   "Arrange &Icons"
      End
   End
   Begin VB.Menu help_mnu 
      Caption         =   "&Help"
      Begin VB.Menu mnu_ContentsIndex 
         Caption         =   "&Contents and Index"
      End
      Begin VB.Menu mnu_tipofday 
         Caption         =   "&Tip of the  Day"
      End
      Begin VB.Menu mnu_perlresources 
         Caption         =   "&CGI Resources"
         Begin VB.Menu mnu_cgimadeeasy 
            Caption         =   "CGI Made Easy"
         End
         Begin VB.Menu mnu_cgiresourceindex 
            Caption         =   "CGI Resource Index"
         End
         Begin VB.Menu mnu_cgiscape 
            Caption         =   "CGI Scape"
         End
         Begin VB.Menu mnu_cgiworks 
            Caption         =   "CGI Works"
         End
         Begin VB.Menu mnu_cgiworld 
            Caption         =   "CGI World"
         End
         Begin VB.Menu mnu_mattsriptarchive 
            Caption         =   "Matt's Script Archive"
         End
         Begin VB.Menu mnu_scriptlocker 
            Caption         =   "Script Locker"
         End
         Begin VB.Menu mnu_thecgicollection 
            Caption         =   "The CGI Collection"
         End
         Begin VB.Menu mnu_cgidirectory 
            Caption         =   "The CGI Directory"
         End
      End
      Begin VB.Menu mnu_perlreference 
         Caption         =   "&Perl References"
         Begin VB.Menu mnu_wwwperlcom 
            Caption         =   "www.perl.com"
         End
         Begin VB.Menu mnu_perlline1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_effectiveperl 
            Caption         =   "Effective Perl"
         End
         Begin VB.Menu mnu_motherofperl 
            Caption         =   "Mother of Perl"
         End
         Begin VB.Menu mnu_perlmongers 
            Caption         =   "Perl Mongers"
         End
         Begin VB.Menu mnu_perlreferencelink 
            Caption         =   "Perl Reference"
         End
         Begin VB.Menu mnu_theperljournal 
            Caption         =   "The Perl Journal"
         End
      End
      Begin VB.Menu mnu_helpline1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_technicalsupport 
         Caption         =   "&Technical Support"
      End
      Begin VB.Menu mnu_elucidonline 
         Caption         =   "&Elucid Software Online"
         Begin VB.Menu mnu_homepage 
            Caption         =   "&Homepage"
         End
         Begin VB.Menu mnu_elucidonlineline1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_software 
            Caption         =   "&Software"
         End
         Begin VB.Menu mnu_developers 
            Caption         =   "&Developers"
         End
         Begin VB.Menu mnu_tools 
            Caption         =   "&Tools"
         End
         Begin VB.Menu mnu_otherservices 
            Caption         =   "&Other Services"
         End
      End
      Begin VB.Menu fd 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################################################################
'#                            Production Of Elucid Software Inc.                       #
'#                                                                                     #
'#  This has been a production of ES Inc. and is not to be reproduced in any way       #
'#  without permission from ES itself and PB.                                          #
'#                                                                                     #
'#  Programmer: Eric Malamisura                                                        #
'#  Programmer: Paul Beviss                                                            #
'#  Last Modified Date: 11/9/00                                                       #
'#  Webpage: http://elucidsoftware.hypermart.net                                       #
'#                                                                                     #
'#  CgEye - CGI IDE PRODUCTION TOOL                                                    #
'#######################################################################################


Option Explicit
Public txtChanged As TriState

Public Enum TriState
    tsTrue = -1
    tsFalse = 0
    tsnone = -2
End Enum

Dim TextHeigth As Long, fTop As Integer
Dim LineCountChange As Integer
Dim FirstLine As Long
Dim FirstLineNow As Long

Public WindowNumber As Integer

Public bRedoing As Boolean
Public UndoStack As New Collection
Public RedoStack As New Collection
Public lUndoCount As Long
Private Const cFilters As String = "All CGI(*.pl *.cgi)|*.pl;*.cgi|Cgi(*.cgi)|*.cgi|Perl(*.pl)|*.pl|All Files(*.*)|*.*"

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As Long    ' /* allocated by caller, zero terminated by RichEdit */
End Type

Private Const WM_USER = &H400
Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const CB_SETDROPPEDWIDTH = &H160&
Private Const LB_FINDSTRING = &H18F
Private Const LB_ERR = (-1)
Private Const EM_POSFROMCHAR = &HD6&
Private Const EM_LINEFROMCHAR = &HC9


Private strTotal As String
Private strPartial As String
Private lCurrentKeyWordIndex As Long
Private bEditFromCode As Boolean
Private bNoChange As Boolean

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Plugins*************
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
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

Private Function CorrectPath(strPath As String) As String

If Len(strPath) = 3 Then
    CorrectPath = strPath
Else
    CorrectPath = strPath & "\"
End If

End Function

Private Sub GetAvailPlugins()


On Error GoTo handler:

filePlugin.Path = App.Path & "\Plugins"
filePlugin.Pattern = "*.dll"

For a = 0 To filePlugin.ListCount - 1
    strTemp = Left(filePlugin.List(a), InStr(filePlugin.List(a), ".") - 1)
    Set MyObj = CreateObject(strTemp & ".Main")
    Call AddToPluginMenu(MyObj.pluginId, strTemp & ".Main")
    strTemp = MyObj.pluginId & " - " & MyObj.pluginversion & " - " & MyObj.plugindescription
    Set MyObj = Nothing
Next

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

Static PluginIndex As Long

If PluginIndex > 0 Then Load mnupluginoptions(PluginIndex)

With mnupluginoptions(PluginIndex)
   .Enabled = True
   .Caption = strPlugin
   .Tag = strKey
End With

PluginIndex = PluginIndex + 1
mnuplugins.Enabled = True
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



Private Sub mnu_about_Click()
frmAbout.Show 1, mdiMain
End Sub

'End Plugin**************


Private Sub mnu_Bold_Click()
    mdiMain.ActiveForm.InsertTag "<B>", "</B>"
End Sub

Private Sub mnu_CenterTag_Click()
   mdiMain.ActiveForm.InsertTag "<p align=\""center\"" >", "</p>"
End Sub

Private Sub mnu_CommentBlock_Click()
   mdiMain.ActiveForm.CommentBlock
End Sub

Private Sub mnu_delete_Click()
  Dim cUndo As clsUndo
    Set cUndo = New clsUndo
    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sDelText = mdiMain.ActiveForm.txtMain.SelText
    mdiMain.ActiveForm.AddToUndoStack cUndo
    txtMain.SelText = ""
End Sub

Private Sub mnu_documents_Click(Index As Integer)
    mdiMain.NewDocument
    mdiMain.ActiveForm.txtChanged = -2
    mdiMain.ActiveForm.txtMain.LoadFile mnu_documents(Index).Tag, rtfText
    mdiMain.ActiveForm.txtChanged = 0
    mdiMain.ActiveForm.Text1.Text = mnu_documents(Index).Tag
    mdiMain.ActiveForm.Caption = FileGetName(mnu_documents(Index).Tag)
    mdiMain.ListSubs
End Sub

Private Sub mnu_effectiveperl_Click()
   OpenIt "http://www.effectiveperl.com/"
End Sub

Private Sub mnu_find_Click()
frmFind.Show , mdiMain
End Sub

Private Sub mnu_formstoolbar_Click()
  Dim CReg As New CRegister
    Set CReg = New CRegister

    If mdiMain.Picture3.Visible = True Then

        mdiMain.Picture3.Visible = False
        mnu_formstoolbar.Checked = False
        mdiMain.mnu_formstoolbar.Checked = False
        mdiMain.mnu_hidetoolsbar2.Checked = False
      Else
        mdiMain.Picture3.Visible = True
        mnu_formstoolbar.Checked = True
        mdiMain.mnu_formstoolbar.Checked = True
        mdiMain.mnu_hidetoolsbar2.Checked = True
    End If

    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowFormsToolbar", mnu_formstoolbar.Checked
    Set CReg = Nothing
End Sub

Private Sub mnu_gotoline_Click()
   frmGotoLine.Show
End Sub

Private Sub mnu_homepage_Click()
   OpenIt "http://elucidsoftware.hypermart.net"
End Sub

Private Sub mnu_insertfont_Click()
frmFont.Show , mdiMain
End Sub

Private Sub mnu_insertpicture_Click()
frmInsertImage.Show , mdiMain
End Sub

Private Sub mnu_inserturl_Click()
frmInsertURL.Show , mdiMain
End Sub



Private Sub mnu_insform_Click(Index As Integer)
        If mdiMain.ActiveForm Is Nothing Then Exit Sub
        
        Select Case Index
        Case 0
        frminsertforms.Show 0, mdiMain
        frminsertforms.Panel6.Visible = True
        Case 1
        frminsertforms.Show 0, mdiMain
        frminsertforms.Panel4.Visible = True
        Case 2
        frminsertforms.Show 0, mdiMain
        frminsertforms.Panel1.Visible = True
        Case 3
        frminsertforms.Show 0, mdiMain
        frminsertforms.Panel2.Visible = True
        Case 4
        frminsertforms.Show 0, mdiMain
        frminsertforms.Panel5.Visible = True
        Case 5
        frminsertforms.Show 0, mdiMain
        frminsertforms.Panel3.Visible = True
        Case 6
        frminsertforms.Show 0, mdiMain
        Case 7
        frminsertforms.Show 0, mdiMain
        frminsertforms.Panel7.Visible = True
End Select
End Sub

Private Sub mnu_Italic_Click()
    mdiMain.ActiveForm.InsertTag "<I>", "</I>"
End Sub

Private Sub mnu_LeftTag_Click()
    mdiMain.ActiveForm.InsertTag "<p align=\""left\"" >", "</p>"
End Sub

Private Sub mnu_leftwin_Click()
      If mdiMain.leftwin.Visible = True Then
        mdiMain.leftwin.Visible = False
        mnu_leftwin.Checked = False
      Else
        mdiMain.leftwin.Visible = True
        mnu_leftwin.Checked = True
    End If
    
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", mnu_leftwin.Checked
    Set CReg = Nothing

End Sub

Private Sub mnu_linenumbers_Click()
On Error Resume Next
    If mdiMain.ActiveForm.picLines.Visible = False Then
        mdiMain.ActiveForm.picLines.Visible = True
        mnu_linenumbers.Checked = True
        mdiMain.varShowLines = True
      Else
        mdiMain.ActiveForm.picLines.Visible = False
        mnu_linenumbers.Checked = False
        mdiMain.varShowLines = False
    End If
mdiMain.ActiveForm.Form_Resize

  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", mnu_linenumbers.Checked
    Set CReg = Nothing

End Sub

Private Sub mnu_newtemplate_Click()
frmTemplate.Show 0, mdiMain
End Sub

Private Sub mnu_printline_Click()
mdiMain.ActiveForm.InsertTag "Print ", ""
End Sub

Private Sub mnu_tables_Click()
frmTables.Show , mdiMain
End Sub

Private Sub mnu_tilehorizontally_Click()
    mdiMain.Arrange vbHorizontal
End Sub

Private Sub mnu_tilevertically_Click()
    mdiMain.Arrange vbVertical
End Sub
Private Sub mnu_arrangeicons_Click()
    mdiMain.Arrange vbArrangeIcons
End Sub

Private Sub mnu_cascade_Click()
    mdiMain.Arrange vbCascade
End Sub
Private Sub mnu_motherofperl_Click()
    OpenIt "http://www.webreference.com/perl/"
End Sub

Private Sub mnu_redo_Click()
    mdiMain.cmdRedo_Click
End Sub

Private Sub mnu_perlmongers_Click()
OpenIt "http://www.perl.org/"
End Sub

Private Sub mnu_perlreferencelink_Click()
OpenIt "http://www.perl.com/reference/query.cgi?cgi"
End Sub

Private Sub mnu_RightTag_Click()
   
     mdiMain.ActiveForm.InsertTag "<p align=\""right\"" >", "</p>"
End Sub

Private Sub mnu_scriptlocker_Click()
OpenIt "http://www.scriptlocker.com/"
End Sub

Private Sub mnu_settings_Click()
    frmSettings.Show , mdiMain
End Sub

Private Sub mnu_software_Click()
OpenIt "http://elucidsoftware.hypermart.net/software.htm"
End Sub

Private Sub mnu_statusbar_Click()
   ' If mdiMain.picStatusBar.Visible = True Then
    '    mdiMain.picStatusBar.Visible = False
     '   mnu_statusbar.Checked = False
     ' Else
      '  mdiMain.picStatusBar.Visible = True
        mnu_statusbar.Checked = True
   ' End If
    
    
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", mnu_statusbar.Checked
    Set CReg = Nothing
End Sub

Private Sub mnu_tipofday_Click()
SaveSetting App.EXEName, "Options", "Show Tips at Startup", 1
    frmTip.Show vbModal, mdiMain
End Sub

Private Sub mnu_toolbar_Click()
  Dim CReg As New CRegister
    Set CReg = New CRegister

    If mdiMain.pictoolbar.Visible = True Then

        mdiMain.pictoolbar.Visible = False
        mnu_toolbar.Checked = False
        mdiMain.mnu_toolbar.Checked = False
        mdiMain.mnu_hidetoolbar.Checked = False
      Else
        mdiMain.pictoolbar.Visible = True
        mnu_toolbar.Checked = True
        mdiMain.mnu_toolbar.Checked = True
        mdiMain.mnu_hidetoolbar.Checked = True
    End If

    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", mnu_toolbar.Checked
    Set CReg = Nothing
End Sub

Private Sub mnu_uncommentblock_Click()
    mdiMain.ActiveForm.UncommentBlock
End Sub

Private Sub mnu_Underline_Click()
    mdiMain.ActiveForm.InsertTag "<U>", "</U>"
End Sub

Private Sub mnu_wwwperlcom_Click()
OpenIt "http://www.perl.com/pub"
End Sub



Private Sub mnuEditCopy_Click()
    SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_COPY, 0, 0
End Sub

Private Sub mnuEditCut_Click()
  Dim cUndo As clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sDelText = mdiMain.ActiveForm.txtMain.SelText
    SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_CUT, 0, 0
    mdiMain.ActiveForm.AddToUndoStack cUndo
End Sub

Private Sub mnuEditDSelectAll_Click()
 With mdiMain.ActiveForm.txtMain
        .SetFocus
        .SelStart = 0
        .SelLength = Len(mdiMain.ActiveForm.txtMain.Text)
    End With
End Sub

Private Sub mnuEditPaste_Click()
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_PASTE, 0, 0
    cUndo.sAddText = Clipboard.GetText(vbCFText)
    mdiMain.ActiveForm.AddToUndoStack cUndo
End Sub

Private Sub mnuEditUndo_Click()
mdiMain.cmdUndo_Click
End Sub


Private Sub mnu_cgidirectory_Click()
OpenIt "http://www.cgidir.com/"
End Sub

Private Sub mnu_cgimadeeasy_Click()
OpenIt "http://www.jmarshall.com/easy/cgi/"
End Sub

Private Sub mnu_cgiresourceindex_Click()
OpenIt "http://www.cgi-resources.com/"
End Sub

Private Sub mnu_cgiscape_Click()
OpenIt "http://www.cgi.tj/"
End Sub

Private Sub mnu_cgiworks_Click()
OpenIt "http://www.cgi-works.net/"
End Sub

Private Sub mnu_cgiworld_Click()
OpenIt "http://www.cgi-world.com/"
End Sub

Private Sub mnu_thecgicollection_Click()
OpenIt "http://www.itm.com/cgicollection/"
End Sub

Private Sub mnu_theperljournal_Click()
OpenIt "http://www.itknowledge.com/tpj/"
End Sub
Public Property Get CanPaste() As Boolean
    CanPaste = SendMessageLong(txtMain.hwnd, EM_CANPASTE, 0, 0)
End Property
Public Sub UpdateStatus()
    mdiMain.cmdRedo.Enabled = Not (RedoStack.Count = 0)
    mdiMain.cboRedo.Enabled = Not (RedoStack.Count = 0)
    mdiMain.cmdUndo.Enabled = Not (UndoStack.Count = 0)
    mdiMain.cboUndo.Enabled = Not (UndoStack.Count = 0)
    
End Sub
Public Function InsertTag(Tag1 As String, Tag2 As String, Optional PrintString As Boolean = False)
    On Error Resume Next
    Dim cUndo As New clsUndo

      cUndo.lStart = txtMain.SelStart
    Dim Buf1 As Integer, Buf2 As Integer
      Buf1 = txtMain.SelStart
      cUndo.lStart = txtMain.SelStart
      If txtMain.SelLength > 0 Then
          cUndo.sDelText = txtMain.SelText
          Buf2 = txtMain.SelLength
          txtMain.SelLength = 0
          txtMain.SelStart = Buf1
          
          If PrintString = True Then
              txtMain.SelText = "print " & """" & Tag1
              txtMain.SelStart = Buf1 + Buf2 + Len(Tag1) + 7
              txtMain.SelText = Tag2 & """"
              txtMain.SelStart = Buf1
              txtMain.SelLength = Buf2 + Len(Tag1) + Len(Tag2) + 8
            Else
              txtMain.SelText = Tag1
              txtMain.SelStart = Buf1 + Buf2 + Len(Tag1)
              txtMain.SelText = Tag2
              txtMain.SelStart = Buf1
              txtMain.SelLength = Buf2 + Len(Tag1) + Len(Tag2)
          End If
          cUndo.sAddText = txtMain.SelText
        Else
        
          If PrintString = True Then
              txtMain.SelText = "print " & """" & Tag1 & Tag2 & """"
              txtMain.SelStart = Buf1 + Len(Tag1) + 7
              cUndo.sAddText = "print " & """" & Tag1 & Tag2 & """"
            Else
              txtMain.SelText = Tag1 & Tag2
              txtMain.SelStart = Buf1 + Len(Tag1)
              cUndo.sAddText = Tag1 & Tag2
          End If
        
      End If
      AddToUndoStack cUndo
      txtMain.SetFocus
End Function

Public Sub InsertFont(Face As String, Size As Integer, Style As String, Color As String)
  Dim sBuf1 As String
  Dim sBuf2 As String

    If Size = 0 And Face = "" Then
        sBuf1 = "<font>"

      ElseIf Size > 0 And Face = "" And Color = "" Then
        sBuf1 = "<font size=\" & """" & Size & "\" & """" & ">"

      ElseIf Size > 0 And Face = "" And Not Color = "" Then
        sBuf1 = "<font size=\" & """" & Size & "\" & """" & " color=\" & """" & Color & "\" & """" & ">"
      ElseIf Not Face = "" And Size = 0 And Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & ">"
      ElseIf Not Face = "" And Size = 0 And Not Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & " color=\" & """" & Color & "\" & """" & ">"
      ElseIf Not Face = "" And Size > 0 And Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & " size=\" & """" & Size & "\" & """" & ">"
      ElseIf Not Face = "" And Size > 0 And Not Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & " size=\" & """" & Size & "\" & """" & " color=\" & """" & Color & "\" & """" & ">"
    End If

    sBuf2 = "</font>"

    If Style = "Italic" Then
        sBuf1 = sBuf1 & "<I>"
        sBuf2 = "</I></font>"
      ElseIf Style = "Bold" Then
        sBuf1 = sBuf1 & "<B>"
        sBuf2 = "</B></font>"
      ElseIf Style = "Bold Italic" Then
        sBuf1 = sBuf1 & "<B><I>"
        sBuf2 = "</I></B></font>"
    End If

    InsertTag sBuf1, sBuf2
End Sub

Public Sub InsertURL(URL As String, Optional Frame As Boolean, Optional FrameType As String)
    If Frame = True Then
        InsertTag "<a href=\" & """" & URL & "\" & """" & " " & "target=\" & """" & FrameType & "\" & """" & ">", "</a>", True
      Else
        InsertTag "<a href=\" & """" & URL & "\" & """" & ">", "</a>", True
    End If

End Sub
Public Sub InsertImagelink(URL As String, border As Integer, Optional Sizes As Boolean, Optional Width As Integer, Optional Height As Integer)
    If Sizes = False Then
        InsertTag "<img " & "border=\" & """" & border & "\" & """" & " src=\" & """" & URL & "\" & """" & ">", "", False
      Else
        InsertTag "<img " & "border=\" & """" & border & "\" & """" & " src=\" & """" & URL & "\" & """" & " width=\" & """" & Width & "\" & """" & " " & "height=\" & """" & Height & "\" & """" & ">", "", False
    End If
End Sub

Public Sub InsertImage(URL As String, border As Integer, Optional Sizes As Boolean, Optional Width As Integer, Optional Height As Integer)
    If Sizes = False Then
        InsertTag "<img " & "border=\" & """" & border & "\" & """" & " src=\" & """" & URL & "\" & """" & ">", "\n" & Chr(34) & ";" & vbCrLf, True
      Else
        InsertTag "<img " & "border=\" & """" & border & "\" & """" & " src=\" & """" & URL & "\" & """" & " width=\" & """" & Width & "\" & """" & " " & "height=\" & """" & Height & "\" & """" & ">", "\n" & Chr(34) & ";" & vbCrLf, True
    End If
End Sub

Private Sub Form_GotFocus()
    UpdateStatus
    
End Sub

Private Sub Form_Load()
    
    mdiMain.WindowCount = mdiMain.WindowCount + 1
    WindowNumber = mdiMain.WindowCount
    UpdateToolbar
    If txtMain.FileName = "" Then Caption = "Document" & WindowNumber & ".cgi"
    mnu_formstoolbar.Checked = mdiMain.mnu_formstoolbar.Checked
    lUndoCount = mdiMain.varUndoLimit
    SetComboDropDownWidth mdiMain.cboUndo, 110
    SetComboDropDownWidth mdiMain.cboRedo, 110
    'this will be editable
    txtChanged = tsFalse
    TextHeigth = txtMain.Font.Size  '// We need this to find out about the size of font
    UpdateStatus
    Text1.Width = ScaleWidth
    Text2.Left = ScaleWidth - (Text2.Width + 150)
Call GetAvailPlugins
Chars_Lines
End Sub
Public Function CloseDocument() As Boolean
    Select Case txtChanged
      Case tsTrue
  Dim YESNO
        YESNO = MsgBox("The following document: " & vbCrLf & Right(Me.Caption, Len(Me.Caption) - 1) & vbCrLf & vbCrLf & "Has had changes made to it." & vbCrLf & vbCrLf & "Would you like to save the changes?", vbYesNoCancel + vbQuestion, "Save Changes?")
        If YESNO = vbYes Then
            ShowSave
            CloseDocument = False
          ElseIf YESNO = vbNo Then
            CloseDocument = False
          ElseIf YESNO = vbCancel Then
            CloseDocument = True
        
        End If
      Case Else
        CloseDocument = False
    End Select
    
End Function

Private Sub Form_Paint()
    If mdiMain.varShowLines = True Then
        DrawNumbers
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = CloseDocument
End Sub

Public Sub Form_Resize()
On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub 'if you minimize it why do you want to resize?
    If Me.ScaleHeight <= 925 Then Exit Sub

    picLines.Move 0, topbar.Height - 20, 610, Me.ScaleHeight - (940 + topbar.Height)
    txtParameters.Move 0, Me.ScaleHeight - 910, Me.ScaleWidth, 910
     If mdiMain.varShowLines = True Then 'check to see if lines are showing
        picLines.Visible = True
        txtMain.Move picLines.Width - 30, topbar.Height - 20, Me.ScaleWidth - (picLines.Width - 30), Me.ScaleHeight - (940 + topbar.Height)
      Else
        txtMain.Move 0, topbar.Height - 20, Me.ScaleWidth, Me.ScaleHeight - (940 + topbar.Height)
        picLines.Visible = False
    End If
Text2.Left = ScaleWidth - (Text2.Width + 150)
topbar1.Width = Me.ScaleWidth + 30
mdiMain.TabStrip1.Move 0, 0, mdiMain.leftwin.Width, mdiMain.leftwin.Height - 60
End Sub

Private Sub Form_Unload(Cancel As Integer)


'disable undo/redo because form in focus is no longer valid..
'mdiMain.Timer1.Enabled = True

UpdateStatus

mdiMain.cmdUndo.Enabled = False
mdiMain.cmdRedo.Enabled = False
mdiMain.cboUndo.Enabled = False
mdiMain.cboRedo.Enabled = False

mdiMain.WindowCount = mdiMain.WindowCount - 1 'subtract the count for the windows
UpdateToolbar
End Sub

Private Sub mnu_revert_Click()
  Dim YESNO As String
    YESNO = MsgBox("Are you sure you wish to revert back to last saved version of this file?" & vbCrLf & vbCrLf & "This will cause all unsaved changes to be lost!", vbYesNo, "Are you sure?")
    
    If YESNO = vbYes Then
        
        mdiMain.ActiveForm.txtMain.LoadFile mdiMain.ActiveForm.txtMain.FileName, rtfText
    End If

End Sub

Private Sub mnuFileClose_Click()
Unload mdiMain.ActiveForm
End Sub

Private Sub mnuFileExit_Click()
Unload mdiMain
End Sub

Private Sub mnuFileNew_Click()
mdiMain.NewDocument
End Sub

Private Sub mnuFileOpen_Click()
       If mdiMain.WindowCount > 9 Then
       MsgBox "You Must Close Some Windows First"
      Else
  Dim CmdDlg As New cCommonDialog
    Set CmdDlg = New cCommonDialog
    CmdDlg.Filter = cFilters
    CmdDlg.DialogTitle = "Open Script"
    CmdDlg.hwnd = Me.hwnd
    CmdDlg.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.ShowOpen
    If CmdDlg.FileName = "" Then Exit Sub
    If mdiMain.ActiveForm Is Nothing Then
        mdiMain.NewDocument
      ElseIf mdiMain.ActiveForm.txtChanged = True Or Len(mdiMain.ActiveForm.txtMain.Text) > 0 Then
        mdiMain.NewDocument
    End If
    mdiMain.ActiveForm.txtChanged = tsnone
    mdiMain.ActiveForm.Caption = CmdDlg.FileTitle
    mdiMain.ActiveForm.Text1.Text = CmdDlg.FileName
    txtMain.LoadFile CmdDlg.FileName, rtfText
    Colorize mdiMain.ActiveForm.txtMain, &H8000&, &HFF0000, &H80&, True
    mdiMain.ActiveForm.Caption = FileGetName(CmdDlg.FileName)
    mdiMain.ActiveForm.txtChanged = tsFalse
    mdiMain.AddRecentList CmdDlg.FileName
    Set CmdDlg = Nothing
End If
mdiMain.ListSubs
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    Dim c As New cCommonDialog
      With c
          .DialogTitle = "Choose Printer"
          .hwnd = Me.hwnd
          .PrinterDefault = True
          .Object = Printer
          .ShowPrinter
          mdiMain.ActiveForm.txtMain.SelPrint Printer.hdc
      End With
End Sub

Private Sub mnuFilePrintSetup_Click()
    On Error Resume Next
    Dim c As New cCommonDialog
      With c
          .DialogTitle = "Choose Printer"
          .hwnd = Me.hwnd
          .PrinterDefault = True
          .Object = Printer
          .flags = PD_PRINTSETUP
          .ShowPrinter
      End With
End Sub

Public Sub mnuFileSave_Click()
Dim hfile As Long
Dim strFilename As String
If mdiMain.ActiveForm.Text1.Text = "New Document" Then
Call ShowSave
Else
mdiMain.ActiveForm.Text1.Text = FileGetName(mdiMain.ActiveForm.txtMain.FileName)

If mdiMain.ActiveForm.txtMain.FileName = "" Then

        If Left(mdiMain.ActiveForm.Caption, 1) = "*" Then
            strFilename = Right(mdiMain.ActiveForm.Caption, Len(mdiMain.ActiveForm.Caption) - 1)
          Else
            strFilename = mdiMain.ActiveForm.Caption
        End If
      Else
        strFilename = mdiMain.ActiveForm.txtMain.FileName
    End If
hfile = FreeFile
Open strFilename For Output As hfile
    Print #hfile, mdiMain.ActiveForm.txtMain.Text
Close
    mdiMain.ActiveForm.Caption = FileGetName(mdiMain.ActiveForm.txtMain.FileName)
    mdiMain.ActiveForm.txtChanged = tsFalse
    If mdiMain.varClearUndo = True Then mdiMain.ActiveForm.ClearUndoRedo
End If
End Sub

Public Sub mnuFileSaveAs_Click()
ShowSave
End Sub
Private Sub ShowSave()
  Dim CmdDlg As New cCommonDialog
    Set CmdDlg = New cCommonDialog
    CmdDlg.Filter = cFilters

    If mdiMain.ActiveForm.txtMain.FileName = "" Then

        If Left(mdiMain.ActiveForm.Caption, 1) = "*" Then
            CmdDlg.FileName = Right(mdiMain.ActiveForm.Caption, Len(mdiMain.ActiveForm.Caption) - 1)
          Else
            CmdDlg.FileName = mdiMain.ActiveForm.Caption
        End If
      Else
        CmdDlg.FileName = mdiMain.ActiveForm.txtMain.FileName
    End If
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.DialogTitle = "Save Script"
    CmdDlg.flags = OFN_OVERWRITEPROMPT
    CmdDlg.hwnd = Me.hwnd
    CmdDlg.ShowSave
    If CmdDlg.FileName = "" Then Exit Sub
'    ActiveForm.txtMain.Text = Replace(ActiveForm.txtMain.Text, vbCrLf, vbCr)
    Dim hfile As Long
    hfile = FreeFile

    Open CmdDlg.FileName For Output As hfile
    Print #hfile, mdiMain.ActiveForm.txtMain.Text
    Close

    mdiMain.ActiveForm.Caption = FileGetName(CmdDlg.FileName)
    'mdiMain.AddRecentList CmdDlg.FileName
    mdiMain.ActiveForm.txtChanged = tsFalse
    If mdiMain.varClearUndo = True Then mdiMain.ActiveForm.ClearUndoRedo
    Set CmdDlg = Nothing
End Sub

Private Sub mnupluginoptions_Click(Index As Integer)
On Error GoTo handler:
'run the plugin..
strTemp = mnupluginoptions(Index).Tag

Set MyObj = CreateObject(strTemp)

Do While MyObj.pluginstart = 1
    DoEvents
Loop

'get the information back from the plugin dll..
If MyObj.strReturn <> "" Then
    'MsgBox "Returned from plugin : " & MyObj.strReturn
    mdiMain.ActiveForm.InsertTag MyObj.strReturn, ""
   
End If

Set MyObj = Nothing

Exit Sub

handler:
MsgBox Err.Number & " - " & Err.Description, vbExclamation

End Sub

Private Sub Text2_Change()
Text2.Left = ScaleWidth - (Text2.Width + 150)
End Sub

Private Sub Timer1_Timer()
'// Get first visible line in rtfText
    DoEvents
    FirstLine = SendMessage(txtMain.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary
    DoEvents
    If Not FirstLineNow = FirstLine Then DrawNumbers '// I can't hook to a scrollbar so I used a sucker-timer
    DoEvents

End Sub

Private Sub txtMain_Change()
  Dim LineCount As Long
    If mdiMain.ActiveForm.txtChanged = tsFalse Then
        Me.Caption = "*" & Me.Caption
    End If
    mdiMain.ActiveForm.txtChanged = tsTrue

    '// Get number of lines in Rtftext
    LineCount = SendMessage(txtMain.hwnd, EM_GETLINECOUNT, 0&, 0&)
    LineCount = LineCount - 1  '// Change start from 0 to 1

    If LineCount = LineCountChange Then
        GoTo skip:    '// Line count is still the same
      Else
        DrawNumbers '// new Line count is required
    End If
    
skip:

    If bRedoing Then
        bRedoing = False
        ClearStack RedoStack
    End If

    UpdateStatus
    UpdateCopyPaste
Chars_Lines
End Sub

Private Sub txtMain_Click()
'txtParameters.Visible = False
UpdateStatus
End Sub

Private Sub txtMain_GotFocus()
    On Error Resume Next
    Dim Cntrl As Control
      For Each Cntrl In Controls
          Cntrl.TabStop = False
      Next Cntrl
      UpdateStatus
      UpdateCopyPaste
      
End Sub

Sub DrawNumbers()
  Dim LineCount As Long '// How many lines in total
  Dim i As Long      '// Just an integer
  Dim TempBuf As String
  Static WidthCount As Integer
'// Get number of lines in Rtftext
    LineCount = SendMessage(txtMain.hwnd, EM_GETLINECOUNT, 0&, 0&)
    LineCount = LineCount - 1  '// Change start from 0 to 1

    '// Same lines ?
    LineCountChange = LineCount

    '// Get first visible line in rtfText
    FirstLine = SendMessage(txtMain.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary

    picLines.Cls '// Clear the PicLines
    picLines.CurrentY = 40  '// Move the .top text by 40 twips

    '// Print the number of each line on a picture
    For i = 0 To LineCount - FirstLine
        picLines.CurrentY = picLines.CurrentY + 7.49 '// Where on Y
        picLines.CurrentX = 20 '-2                   '// Where on X
        picLines.Print i + FirstLine + 1             '// print the number
    Next
    picLines.Refresh
    'LineCountChange = LineCount '// Remember the last line count
    FirstLineNow = FirstLine     '// Is the first visible line still the same ?
End Sub
Public Sub GotoLine(LineNum As Long, Highlight As Boolean)
    On Error GoTo done:
  Dim Temp As Integer
  Dim Num As Integer
  Dim Pos  As Integer
  Dim LastPos As Integer
  Dim Cut As Integer
    If LineNum = 0 Then Exit Sub
    Pos = 1
    Num = 1
    Temp = 0
    Do
        LastPos = Temp
        Temp = InStr(Pos, txtMain.Text, vbLf)
        If Temp = 0 Then GoTo redo:
        If Temp >= 1 Then
            Num = Num + 1
            Pos = Temp + 2
        End If
    
    Loop Until Num >= LineNum

    Cut = 1

redo:
    If Temp = 0 Then
        LastPos = 0
        Temp = Len(txtMain.Text)
        Cut = 0
    End If

    If LineNum = 1 Then
        Temp = 0
        LastPos = InStr(1, txtMain.Text, vbLf)
        If LastPos = 0 Then
            LastPos = Len(txtMain.Text)
        End If

        Cut = 0
    End If

    txtMain.SelStart = Temp
    If Highlight = True Then txtMain.SelLength = LastPos - Cut
    txtMain.SetFocus
done:
End Sub

Public Function GetUndoText(ModifyType As ModifyTypes) As String
    Select Case ModifyType
      Case DeleteText
        GetUndoText = "Delete Text"
      Case AddText
        GetUndoText = "Add Text"
      Case ReplaceText
        GetUndoText = "Replace Text"
      Case PasteText
        GetUndoText = "Paste Text"
      Case CutText
        GetUndoText = "Cut Text"
    End Select
End Function
Public Sub AddToUndoStack(cUndo As clsUndo)
    If UndoStack.Count = lUndoCount Then
        UndoStack.Remove (1)
    End If
    UndoStack.Add cUndo
    UpdateStatus
End Sub
Private Sub ClearStack(Stack As Collection)
    On Error Resume Next
    Dim i As Long
      For i = 1 To Stack.Count Step 1
          Stack.Remove (i)
      Next
End Sub
Public Sub ClearUndoRedo()
  Dim i As Long

    For i = 1 To UndoStack.Count
        UndoStack.Remove (1)
    Next i

    i = 1

    For i = 1 To RedoStack.Count
        RedoStack.Remove (1)
    Next i

    UpdateStatus
End Sub
Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cUndo As New clsUndo
  Dim lAmount As Long
  Dim lOldPos As Long
Chars_Lines
    If IsMoveKey(KeyCode) Then Exit Sub

    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If KeyCode = vbKeyBack Then
            lAmount = -1
          Else
            lAmount = 0
        End If
        With cUndo
            If txtMain.SelLength = 0 Then
                '// there is no text being deleted
                If lAmount = -1 And txtMain.SelStart = 0 Then
                    '// we aren't going anywhere!
                    GoTo exitundo
                End If
                '// set the start pos
                .lStart = IIf(txtMain.SelStart = 0, 0, txtMain.SelStart + lAmount)
                '// see what we are going to delete
                .sDelText = TextInRange(txtMain.SelStart + lAmount, 1)
                
                '// if there is part of vbCrLf selected
                '// set the length to 2 instead
                If InStr(1, Chr(10) & Chr(13), .sDelText) Then
                    .sDelText = vbCrLf
                    If .sDelText = Chr(10) Then
                        '// deleting end of CrLf
                        .lStart = .lStart - 1
                    End If
                End If
              Else
                '// save the text that is being deleted
                .lStart = txtMain.SelStart
                .sDelText = txtMain.SelText
            End If
            .ModifyType = DeleteText
            AddToUndoStack cUndo
exitundo:
        End With
    End If
    
    If Shift = vbCtrlMask And KeyCode <> vbKeyControl Then
        With cUndo
            Select Case KeyCode
              Case vbKeyV
                '// add the pasted text to the Undo stack
                .lStart = txtMain.SelStart
                .sAddText = Clipboard.GetText(vbCFText)
                .sDelText = txtMain.SelText
                .ModifyType = PasteText
                AddToUndoStack cUndo
                txtMain.SelText = .sAddText
                KeyCode = 0
              Case vbKeyX
                '// cut
                .lStart = txtMain.SelStart
                .sDelText = txtMain.SelText
                
                .ModifyType = CutText
                AddToUndoStack cUndo
              Case vbKeyZ
                mdiMain.cmdUndo_Click
                KeyCode = 0
              Case vbKeyY
                mdiMain.cmdRedo_Click
                KeyCode = 0
            End Select
        End With
    End If


End Sub
    
Public Property Get TextInRange(ByVal lStart As Long, ByVal lLen As Long)
  Dim tR As TEXTRANGE
  Dim lR As Long
  Dim sText As String
  Dim b() As Byte
  Dim lEnd As Long

    lEnd = lStart + lLen
    
    tR.chrg.cpMin = lStart
    tR.chrg.cpMax = lEnd
    
    sText = String$(lEnd - lStart + 1, 0)
    b = StrConv(sText, vbFromUnicode)
    ' VB won't do the terminating null for you!
    ReDim Preserve b(0 To UBound(b) + 1) As Byte
    b(UBound(b)) = 0
    tR.lpstrText = VarPtr(b(0))
    
    lR = SendMessage(txtMain.hwnd, EM_GETTEXTRANGE, 0, tR)
    If (lR > 0) Then
        sText = StrConv(b, vbUnicode)
        TextInRange = Left$(sText, lR)
    End If
End Property

Private Function IsMoveKey(KeyCode As Integer) As Boolean
    Select Case KeyCode
      Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd, vbKeyShift
        IsMoveKey = True
    End Select
End Function

Private Sub txtMain_KeyPress(KeyAscii As Integer)
  Dim cUndo As New clsUndo
    
    If KeyAscii = vbKeyBack Then
        '// ignore. Note that the Delete key does not trigger this event
      ElseIf KeyAscii >= 32 Or KeyAscii = 13 Then
        '// ignore keycodes under 32
        With cUndo
            .lStart = txtMain.SelStart
            If KeyAscii = 13 Then
                KeyAscii = 0
                Call CheckSyntax
                KeyAscii = 13
                .sAddText = vbCrLf
              Else
                .sAddText = Chr(KeyAscii)
            End If
            .sDelText = txtMain.SelText
            .ModifyType = IIf(.sDelText = "", AddText, ReplaceText)
        End With
        AddToUndoStack cUndo
    End If
    Set cUndo = Nothing


'    If KeyAscii = 13 Then
'        mdiMain.ListSubs
'    End If

'If KeyAscii = vbKeyBack Then ShowParam False
   
End Sub
Public Sub CommentBlock()
  Dim Buffer As Variant
  Dim cUndo As New clsUndo

    If txtMain.SelLength = 0 Then
        cUndo.lStart = txtMain.SelStart
        txtMain.SelText = "#"
        cUndo.sAddText = "#"
        AddToUndoStack cUndo
      Else
        cUndo.lStart = txtMain.SelStart
        cUndo.sDelText = txtMain.SelText
        Buffer = txtMain.SelText
        Buffer = "#" & Buffer
        Buffer = Replace(Left(Buffer, Len(Buffer)), vbLf, vbLf & "#")
        txtMain.SelText = Buffer
        cUndo.sAddText = Buffer
        AddToUndoStack cUndo
    End If

End Sub

Public Sub UncommentBlock()
  Dim Buffer As Variant
  Dim FirstLineBuffer As Variant
  Dim cUndo As New clsUndo
  Dim Proceed As Boolean
  Dim i As Integer
    cUndo.lStart = txtMain.SelStart
    cUndo.sDelText = txtMain.SelText
    Buffer = txtMain.SelText
    If InStr(Buffer, vbLf) Then
        FirstLineBuffer = Left(Buffer, InStr(Buffer, vbLf))
      Else
        FirstLineBuffer = Buffer
    End If
    If InStr(FirstLineBuffer, "#") Then
        For i = 1 To Len(FirstLineBuffer)
            If Mid(FirstLineBuffer, i, 1) = "" Or Mid(FirstLineBuffer, i, 1) = " " Or Mid(FirstLineBuffer, i, 1) = "#" Then
                Proceed = True
                Exit For
              Else
                Proceed = False
            End If
        Next i
        If Proceed = True Then
            Buffer = Right(Buffer, Len(Buffer) - InStr(FirstLineBuffer, "#"))
        End If
    End If
skiptrim:
    If InStr(Buffer, vbLf & "#") Then
        Buffer = Replace(Left(Buffer, Len(Buffer)), vbLf & "#", vbLf)
    End If
    txtMain.SelText = Buffer
    cUndo.sAddText = Buffer
    AddToUndoStack cUndo
End Sub
Public Sub UpdateCopyPaste()
If mdiMain.ActiveForm Is Nothing Then Exit Sub

If CanPaste = True Then
        mdiMain.Toolbar1.Buttons.Item(8).Enabled = True
        mdiMain.ActiveForm.mnuEditPaste.Enabled = True
    Else
        mdiMain.Toolbar1.Buttons.Item(8).Enabled = False
        mdiMain.ActiveForm.mnuEditPaste.Enabled = False
    End If

    If Len(txtMain.SelText) > 0 Then
       mdiMain.Toolbar1.Buttons.Item(6).Enabled = True
       mdiMain.Toolbar1.Buttons.Item(7).Enabled = True
       mdiMain.Toolbar1.Buttons.Item(9).Enabled = True
    Else
       mdiMain.Toolbar1.Buttons.Item(6).Enabled = False
       mdiMain.Toolbar1.Buttons.Item(7).Enabled = False
       mdiMain.Toolbar1.Buttons.Item(9).Enabled = False
    End If


    If mdiMain.ActiveForm.RedoStack.Count > 1 Then mdiMain.ActiveForm.mnu_redo.Enabled = True
    If mdiMain.ActiveForm.UndoStack.Count > 1 Then mdiMain.ActiveForm.mnuEditUndo.Enabled = True

    If Len(txtMain.SelText) > 0 Then
        mdiMain.ActiveForm.mnuEditCut.Enabled = True
        mdiMain.ActiveForm.mnuEditCopy.Enabled = True
        mdiMain.ActiveForm.mnu_delete.Enabled = True
      Else
        mdiMain.ActiveForm.mnuEditCut.Enabled = False
        mdiMain.ActiveForm.mnuEditCopy.Enabled = False
        mdiMain.ActiveForm.mnu_delete.Enabled = False
    End If

    If Len(txtMain.Text) > 0 Then
        mdiMain.ActiveForm.mnu_gotoline.Enabled = True
        mnu_find.Enabled = True
      '  mnu_findnext.Enabled = True
       ' mnu_replace.Enabled = True
      '  mnu_replacenext.Enabled = True
    End If
End Sub


Private Sub txtMain_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then Call CheckKeyWord
If KeyCode = vbKeyBack Then Call CheckKeyWord


'If KeyCode = vbKeyReturn Then ShowParam False
'If KeyCode = 219 Then ShowParam False 'if { pressed

End Sub

Private Sub txtMain_LostFocus()
    UpdateCopyPaste
End Sub

Private Sub txtMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then

  Me.PopupMenu Me.mnuEdit
 End If
    Chars_Lines
    UpdateCopyPaste
End Sub

Private Sub txtMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateCopyPaste
End Sub
Public Sub GetPosFromChar(ByVal lIndex As Long, ByRef xPixels As Long, ByRef YPixels As Long)
Dim lxy As Long
lxy = SendMessageLong(txtMain.hwnd, EM_POSFROMCHAR, lIndex, 0)
xPixels = (lxy And &HFFFF&)
YPixels = (lxy \ &H10000) And &HFFFF&
xPixels = ScaleX(xPixels, vbPixels, vbTwips)    ' From pixels to twips
YPixels = ScaleY(YPixels, vbPixels, vbTwips)
End Sub
Private Sub ShowParam(ShowIt As Boolean)

End Sub

Public Function GetWorkingLine() As Variant
Dim curPos As Integer
Dim ReturnBefore As Boolean
Dim ReturnAfter As Boolean

ReturnBefore = False
ReturnAfter = False

If txtMain.SelStart > 0 Then
curPos = txtMain.SelStart
Else
curPos = 1
End If
If txtMain.Text = "" Then Exit Function

'Get some info first
If InStr(curPos, txtMain.Text, vbLf) Then 'Need to figure out if return after line
    ReturnAfter = True
Else
    ReturnAfter = False
End If

If InStrRev(txtMain.Text, vbLf, curPos) Then 'Need to figure out if return before line
    ReturnBefore = True
Else
    ReturnBefore = False
End If

'Now we do the real work..
If ReturnBefore = False And ReturnAfter = False Then 'Only one line no others...
    GetWorkingLine = txtMain.Text

End If

If ReturnBefore = True And ReturnAfter = False Then 'Last line in text
    GetWorkingLine = Mid(txtMain.Text, InStrRev(txtMain.Text, vbLf, curPos), Len(txtMain) - InStrRev(txtMain.Text, vbLf, curPos))

End If

If ReturnBefore = False And ReturnAfter = True Then 'first line in text

    GetWorkingLine = Mid(txtMain.Text, 1, InStr(txtMain.Text, vbLf))
End If

If ReturnBefore = True And ReturnAfter = True Then 'line in middle of the text

Dim a As Integer
Dim b As Integer

a = InStrRev(txtMain.Text, vbLf, curPos)
b = InStr(curPos, txtMain.Text, vbLf)
If b - a > 2 Then
GetWorkingLine = Mid(txtMain.Text, a + 1, b - a - 2)
End If
End If

GetWorkingLine = Replace(GetWorkingLine, vbTab, "")
GetWorkingLine = Replace(GetWorkingLine, vbLf, "")
GetWorkingLine = LTrim(GetWorkingLine)


End Function

Public Function IsString() As Boolean
On Error Resume Next
'this detects if its a string or not..by seeing if the text is in " " or ' '

If GetWorkingLine = "" Then Exit Function
IsString = False
If InStrRev(GetWorkingLine, """", txtMain.SelStart) Or InStrRev(GetWorkingLine, "'", txtMain.SelStart) Then 'check to see if its a string or not
    IsString = True
'    If InStr(txtMain.SelStart + 1, GetWorkingLine, """") And InStr(txtMain.SelStart + 1, GetWorkingLine, "'") Then
'        IsString = True
'
'    Else
'        IsString = False
'
'    End If
Else
    IsString = False
End If
End Function

Public Sub CheckKeyWord()
'Ummm this is the code for the tooltips...Im firing it off using the spacebar since in
'perl its about the only thing that seperates keywords..



Dim Keyword As String
Dim Found As Integer
Dim strDescription



If Len(GetWorkingLine) < 2 Then Exit Sub
If Trim(GetWorkingLine) = "" Then Exit Sub

If IsString = False Then
    If InStrRev(GetWorkingLine, " ", Len(GetWorkingLine) - 1) Then
        Keyword = Right(GetWorkingLine, InStrRev(GetWorkingLine, " ", Len(GetWorkingLine) - 1))
    Else
        Keyword = Left(GetWorkingLine, InStr(GetWorkingLine, " "))

    End If
    
    Keyword = Replace(Keyword, vbLf, "") 'Make absolute sure nothing foreign gets in it
    Keyword = Replace(Keyword, " ", "")
      Debug.Print Keyword

        Found = SendMessageByString(mdiMain.lstKeys.hwnd, LB_FINDSTRINGEXACT, -1, Keyword)
              
        
        If Found > -1 Then
            strDescription = mdiMain.lstKeyDescription.List(Found)
            txtParameters.SelBold = True
            txtParameters.SelColor = &H800000
            txtParameters.SelText = "SmartTip)> "
            txtParameters.SelText = strDescription & vbCrLf
        Else
            
        End If


End If

End Sub

Private Sub txtParameters_Change()
txtParameters.SelStart = Len(txtParameters)
End Sub

Public Sub UpdateToolbar()
  Dim i As Integer
  Dim a As Integer
If mdiMain.WindowCount > 0 Then
mdiMain.Toolbar4.Enabled = True
Else
mdiMain.Toolbar4.Enabled = False
End If
End Sub
Public Function CheckSyntax() As Boolean
Dim strDescription As String
CheckSyntax = True

If Trim(GetWorkingLine) = "" Then Exit Function

'**Check the ending line syntax
If Not Mid(Trim(GetWorkingLine), Len(GetWorkingLine), 1) = ";" And Not Mid(Trim(GetWorkingLine), Len(GetWorkingLine), 1) = "{" _
And Not Left(Trim(GetWorkingLine), 1) = "#" Then
    CheckSyntax = False
    strDescription = "Incorrect Line Syntax.  Expected ';' Or '{' at line end."
Else
    CheckSyntax = True
End If
'**------------------------------

If CheckSyntax = False Then
    MsgBox strDescription, vbCritical, "Syntax Error"
    txtParameters.SelBold = True
    txtParameters.SelColor = &H80&
    txtParameters.SelText = "Syntax Error)> "
    txtParameters.SelText = strDescription & vbCrLf
End If

End Function

Private Sub Chars_Lines()
Dim Lines, Chars As String
Dim blah() As String
Dim bleh() As String
Dim Curline As String
Dim CurChar, TotalChar As String

Curline = Mid(txtMain.Text, 1, txtMain.SelStart)
blah() = Split(Curline, Chr$(10))
bleh() = Split(txtMain.Text, Chr$(10))

If txtMain.SelStart = 0 Then
CurChar = 0
Curline = 1

If Len(txtMain.Text) = 0 Then
TotalChar = 0
Else
TotalChar = Len(txtMain.Text) - (UBound(bleh) * 2)
End If

Else
CurChar = txtMain.SelStart - (UBound(blah) * 2)
Curline = UBound(blah) + 1
TotalChar = Len(txtMain.Text) - (UBound(bleh) * 2)
End If

Lines = "Line:" & Curline & "/" & SendMessage(txtMain.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
Chars = "Char:" & CurChar & "/" & TotalChar
Text2.Text = Chars & "  " & Lines
End Sub
