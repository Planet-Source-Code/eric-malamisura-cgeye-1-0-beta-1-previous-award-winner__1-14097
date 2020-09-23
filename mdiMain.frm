VERSION 5.00
Object = "{5C4592BE-A01B-11D3-AFAF-BF3F431B043C}#1.0#0"; "Toolbar2.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H008A8A8A&
   Caption         =   "CgEye By Elucid Software (Preview Release 5)"
   ClientHeight    =   7155
   ClientLeft      =   180
   ClientTop       =   735
   ClientWidth     =   10845
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   10845
      TabIndex        =   13
      Top             =   435
      Width           =   10845
      Begin AIFCmp1.asxToolbar asxToolbar1 
         Height          =   375
         Left            =   0
         Top             =   30
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   661
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGap       =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   9
         PlaySounds      =   0   'False
         ButtonStyle1    =   2
         ButtonKey2      =   "RECTANGL"
         ButtonMaskColor2=   12632256
         ButtonPicture2  =   "mdiMain.frx":030A
         ButtonToolTipText2=   "RECTANGL"
         ButtonKey3      =   "BUTTON"
         ButtonMaskColor3=   12632256
         ButtonPicture3  =   "mdiMain.frx":065C
         ButtonToolTipText3=   "BUTTON"
         ButtonKey4      =   "textinput"
         ButtonMaskColor4=   12632256
         ButtonPicture4  =   "mdiMain.frx":09AE
         ButtonToolTipText4=   "textinput"
         ButtonKey5      =   "muilttext"
         ButtonMaskColor5=   12632256
         ButtonPicture5  =   "mdiMain.frx":0DC0
         ButtonToolTipText5=   "muilttext"
         ButtonKey6      =   "check"
         ButtonMaskColor6=   12632256
         ButtonPicture6  =   "mdiMain.frx":1442
         ButtonToolTipText6=   "check"
         ButtonKey7      =   "optionbutton"
         ButtonMaskColor7=   12632256
         ButtonPicture7  =   "mdiMain.frx":17A0
         ButtonToolTipText7=   "optionbutton"
         ButtonKey8      =   "combo"
         ButtonMaskColor8=   12632256
         ButtonPicture8  =   "mdiMain.frx":1C2A
         ButtonToolTipText8=   "combo"
         ButtonKey9      =   "font"
         ButtonMaskColor9=   12632256
         ButtonPicture9  =   "mdiMain.frx":21CC
         ButtonToolTipText9=   "font"
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   -15
         TabIndex        =   14
         Top             =   -90
         Width           =   10815
      End
   End
   Begin VB.PictureBox pictoolbar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   10845
      TabIndex        =   8
      Top             =   0
      Width           =   10845
      Begin VB.Frame Frame1 
         Height          =   120
         Left            =   15
         TabIndex        =   9
         Top             =   -90
         Width           =   10785
      End
      Begin AIFCmp1.asxToolbar Toolbar1 
         Height          =   375
         Left            =   0
         Top             =   45
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   661
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonGap       =   1
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ButtonCount     =   33
         PlaySounds      =   0   'False
         SolidChecked    =   -1  'True
         ShowSeparators  =   -1  'True
         ButtonMaskColor1=   12632256
         ButtonStyle1    =   2
         ButtonKey2      =   "New"
         ButtonMaskColor2=   12632256
         ButtonPicture2  =   "mdiMain.frx":251E
         ButtonToolTipText2=   "New"
         ButtonKey3      =   "Open"
         ButtonMaskColor3=   12632256
         ButtonPicture3  =   "mdiMain.frx":2870
         ButtonToolTipText3=   "Open"
         ButtonKey4      =   "Save"
         ButtonMaskColor4=   12632256
         ButtonPicture4  =   "mdiMain.frx":2BC2
         ButtonToolTipText4=   "Save"
         ButtonStyle5    =   2
         ButtonKey6      =   "CUT"
         ButtonMaskColor6=   12632256
         ButtonPicture6  =   "mdiMain.frx":2F14
         ButtonToolTipText6=   "CUT"
         ButtonKey7      =   "Copy"
         ButtonMaskColor7=   12632256
         ButtonPicture7  =   "mdiMain.frx":3266
         ButtonToolTipText7=   "Copy"
         ButtonKey8      =   "PASTE"
         ButtonMaskColor8=   12632256
         ButtonPicture8  =   "mdiMain.frx":35B8
         ButtonToolTipText8=   "PASTE"
         ButtonKey9      =   "DELETE"
         ButtonMaskColor9=   12632256
         ButtonPicture9  =   "mdiMain.frx":390A
         ButtonToolTipText9=   "DELETE"
         ButtonStyle10   =   2
         ButtonWidth11   =   90
         ButtonStyle11   =   0
         ButtonStyle12   =   2
         ButtonKey13     =   "color"
         ButtonMaskColor13=   12632256
         ButtonPicture13 =   "mdiMain.frx":3C5C
         ButtonToolTipText13=   "color"
         ButtonKey14     =   "link"
         ButtonMaskColor14=   12632256
         ButtonPicture14 =   "mdiMain.frx":402E
         ButtonToolTipText14=   "link"
         ButtonKey15     =   "picture"
         ButtonMaskColor15=   12632256
         ButtonPicture15 =   "mdiMain.frx":4530
         ButtonToolTipText15=   "picture"
         ButtonStyle16   =   2
         ButtonKey17     =   "LFT"
         ButtonMaskColor17=   12632256
         ButtonPicture17 =   "mdiMain.frx":495E
         ButtonToolTipText17=   "LFT"
         ButtonKey18     =   "CNT"
         ButtonMaskColor18=   12632256
         ButtonPicture18 =   "mdiMain.frx":4CB0
         ButtonToolTipText18=   "CNT"
         ButtonKey19     =   "RT"
         ButtonMaskColor19=   12632256
         ButtonPicture19 =   "mdiMain.frx":5002
         ButtonToolTipText19=   "RT"
         ButtonStyle20   =   2
         ButtonWidth21   =   100
         ButtonStyle21   =   0
         ButtonWidth22   =   60
         ButtonStyle22   =   0
         ButtonStyle23   =   2
         ButtonKey24     =   "font"
         ButtonMaskColor24=   12632256
         ButtonPicture24 =   "mdiMain.frx":5354
         ButtonToolTipText24=   "font"
         ButtonKey25     =   "BLD"
         ButtonMaskColor25=   12632256
         ButtonPicture25 =   "mdiMain.frx":56A6
         ButtonToolTipText25=   "BLD"
         ButtonKey26     =   "ITL"
         ButtonMaskColor26=   12632256
         ButtonPicture26 =   "mdiMain.frx":59F8
         ButtonToolTipText26=   "ITL"
         ButtonKey27     =   "UNDRLN"
         ButtonMaskColor27=   12632256
         ButtonPicture27 =   "mdiMain.frx":5D4A
         ButtonToolTipText27=   "UNDRLN"
         ButtonStyle28   =   2
         ButtonKey29     =   "FIND"
         ButtonMaskColor29=   12632256
         ButtonPicture29 =   "mdiMain.frx":609C
         ButtonToolTipText29=   "FIND"
         ButtonStyle30   =   2
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "mdiMain.frx":63EE
            Left            =   8190
            List            =   "mdiMain.frx":640A
            TabIndex        =   18
            Text            =   "1"
            Top             =   30
            Width           =   570
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   6300
            Sorted          =   -1  'True
            TabIndex        =   17
            Text            =   "Times New Roman"
            Top             =   30
            Width           =   1875
         End
         Begin VB.CommandButton cmdUndo 
            Height          =   255
            Left            =   2700
            Picture         =   "mdiMain.frx":6426
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Undo Last Event"
            Top             =   60
            Width           =   315
         End
         Begin VB.ComboBox cboUndo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2670
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "Undo"
            Top             =   30
            Width           =   630
         End
         Begin VB.CommandButton cmdRedo 
            Height          =   255
            Left            =   3390
            MaskColor       =   &H000000C0&
            Picture         =   "mdiMain.frx":6570
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Redo Last Undo"
            Top             =   60
            Width           =   330
         End
         Begin VB.ComboBox cboRedo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   10
            ToolTipText     =   "Redo"
            Top             =   30
            Width           =   630
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   30
         X2              =   9450
         Y1              =   420
         Y2              =   420
      End
   End
   Begin VB.PictureBox leftwin 
      Align           =   3  'Align Left
      Height          =   6000
      Left            =   0
      ScaleHeight     =   5940
      ScaleWidth      =   1890
      TabIndex        =   6
      Top             =   840
      Width           =   1950
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -30
         TabIndex        =   7
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10845
      TabIndex        =   2
      Top             =   6840
      Width           =   10845
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   -15
         TabIndex        =   12
         Top             =   -90
         Width           =   10815
      End
      Begin VB.ListBox lstKeyDescription 
         Height          =   255
         ItemData        =   "mdiMain.frx":66BA
         Left            =   7440
         List            =   "mdiMain.frx":66BC
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
      Begin VB.ListBox lstKeys 
         Height          =   255
         ItemData        =   "mdiMain.frx":66BE
         Left            =   3480
         List            =   "mdiMain.frx":66C0
         TabIndex        =   4
         Top             =   1770
         Width           =   3735
      End
      Begin VB.Label txtStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Thank you for using CgEye By Elucid Software"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   75
         Width           =   7590
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10845
      TabIndex        =   0
      Top             =   840
      Width           =   10845
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10845
      TabIndex        =   1
      Top             =   840
      Width           =   10845
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_newtemplate 
         Caption         =   "&New From Template"
      End
      Begin VB.Menu mnu_Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_line5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_documents 
         Caption         =   "<blank>"
         Index           =   0
      End
      Begin VB.Menu mnu_line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Begin VB.Menu mnu_toolbar 
         Caption         =   "&Common Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_formstoolbar 
         Caption         =   "&Forms Toolbar"
      End
      Begin VB.Menu mnu_statusbar 
         Caption         =   "&Status Bar"
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
   End
   Begin VB.Menu mnu_toolbat_popup 
      Caption         =   "toolbarpopup"
      Visible         =   0   'False
      Begin VB.Menu mnuTop 
         Caption         =   "Align Top"
      End
      Begin VB.Menu mnu_Abottom 
         Caption         =   "Align Bottom"
      End
      Begin VB.Menu mnu_hidetoolbar 
         Caption         =   "&Hide common toolbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnu_toolbat_popup2 
      Caption         =   "toolbarpop"
      Visible         =   0   'False
      Begin VB.Menu mnu_Top 
         Caption         =   "Align Top"
      End
      Begin VB.Menu mnuAbottom 
         Caption         =   "Align Bottom"
      End
      Begin VB.Menu mnu_hidetoolsbar2 
         Caption         =   "&Hide insert toolbar"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################################################################
'#                            Production Of Elucid Software Inc.                       #
'#                                                                                     #
'#  This has been a production of ES Inc. and is not to be reproduced in any way       #
'#  without permission from ES itself.                                                 #
'#                                                                                     #
'#  Programmer: Eric Malamisura                                                        #
'#  Programmer: Paul Beviss                                                            #
'#  Last Modified Date: 10/24/00                                                       #
'#  Webpage: http://elucidsoftware.hypermart.net                                       #
'#                                                                                     #
'#  CgEye - CGI IDE PRODUCTION TOOL                                                    #
'#######################################################################################


Option Explicit
Private Const cFilters As String = "All CGI(*.pl *.cgi)|*.pl;*.cgi|Cgi(*.cgi)|*.cgi|Perl(*.pl)|*.pl|All Files(*.*)|*.*"
' vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy IP", ""
Public WindowCount As Integer
Private Document(255) As frmMain
Private Declare Function GetTickCount Lib "kernel32" () As Long
'These are for the stupid registry shit that pisses me off
Public varUndoLimit As Integer
Public varClearUndo As Boolean
Public varDocuments As Integer
Public varUseRelative As Boolean
Public varDefaultFolder As String
Public varShowLines As Boolean

Private Sub SaveWindowSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window State", Me.WindowState
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window LeftPos", Me.Left
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window TopPos", Me.Top
    Set CReg = Nothing
End Sub

Private Sub LoadWindowSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window State", vbMaximized) = 2 Then
        Me.WindowState = vbMaximized
        Exit Sub
      Else
        Me.WindowState = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window State", vbMaximized)
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window LeftPos", Me.Left) < 0 Then
        Me.Left = 0
      Else
        Me.Left = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window LeftPos", Me.Left)
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window TopPos", Me.Top) < 0 Then
        Me.Top = 0
      Else
        Me.Top = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window TopPos", Me.Top)
    End If
    Set CReg = Nothing
End Sub



Private Sub asxToolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
        
        If ActiveForm Is Nothing Then Exit Sub
        
        Select Case ButtonIndex
        Case 2
        frminsertforms.Show 0, Me
        frminsertforms.Panel6.Visible = True
        Case 3
        frminsertforms.Show 0, Me
        frminsertforms.Panel4.Visible = True
        Case 4
        frminsertforms.Show 0, Me
        frminsertforms.Panel1.Visible = True
        Case 5
        frminsertforms.Show 0, Me
        frminsertforms.Panel2.Visible = True
        Case 6
        frminsertforms.Show 0, Me
        frminsertforms.Panel5.Visible = True
        Case 7
        frminsertforms.Show 0, Me
        frminsertforms.Panel3.Visible = True
        Case 8
        frminsertforms.Show 0, Me
        frminsertforms.Panel4.Visible = True
        Case 9
        frminsertforms.Show 0, Me
        frminsertforms.Panel7.Visible = True
        
        End Select
End Sub

Private Sub asxToolbar1_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
txtStatus.Caption = ButtonKey
End Sub

Private Sub asxToolbar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then

  Me.PopupMenu Me.mnu_toolbat_popup2
 End If
End Sub

Private Sub mnu_Abottom_Click()
 pictoolbar.Align = 2
End Sub

Private Sub mnu_formstoolbar_Click()
  Dim CReg As New CRegister
    Set CReg = New CRegister

    If Picture3.Visible = True Then

        Picture3.Visible = False
        mnu_formstoolbar.Checked = False
        mnu_hidetoolsbar2.Checked = False
        If ActiveForm Is Nothing Then
        Else
        ActiveForm.mnu_formstoolbar.Checked = False
        End If
        
      Else
        Picture3.Visible = True
        mnu_formstoolbar.Checked = True
        mnu_hidetoolsbar2.Checked = True
        If ActiveForm Is Nothing Then
        Else
        ActiveForm.mnu_formstoolbar.Checked = True
        End If
    End If

    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowFormsToolbar", mnu_formstoolbar.Checked
    Set CReg = Nothing
End Sub

Private Sub mnu_hidetoolbar_Click()
mnu_toolbar_Click
End Sub

Private Sub mnu_hidetoolsbar2_Click()
mnu_formstoolbar_Click
End Sub

Private Sub mnu_newtemplate_Click()
Form1.Show 0, Me
End Sub

Private Sub mnu_Top_Click()
Picture3.Align = 1
End Sub

Private Sub mnuAbottom_Click()
Picture3.Align = 2
End Sub

Private Sub mnuTop_Click()
 pictoolbar.Align = 1
End Sub

Private Sub pictoolbar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then

  Me.PopupMenu Me.mnu_toolbat_popup
 End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then

  Me.PopupMenu Me.mnu_toolbat_popup2
 End If
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
txtStatus = "Thank you for using CgEye By Elucid Software"
End Sub



Public Sub Toolbar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)

        If ButtonIndex > 3 Then
        If ActiveForm Is Nothing Then Exit Sub
        End If
        Dim cUndo As clsUndo
        Set cUndo = New clsUndo
    
    Select Case ButtonIndex
      Case 2 'New Button
       NewDocument
      Case 3 'Open Button
        mnu_Open_Click
      Case 4 'Save Button
        If ActiveForm.txtChanged = -1 Then
            If Len(ActiveForm.txtMain.FileName) > 0 Then
                mdiMain.ActiveForm.mnuFileSave_Click
              Else
                mdiMain.ActiveForm.mnuFileSaveAs_Click
            End If
        End If
      Case 6 'Cut Button
        cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
        cUndo.sDelText = mdiMain.ActiveForm.txtMain.SelText
        SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_CUT, 0, 0
        mdiMain.ActiveForm.AddToUndoStack cUndo
      Case 7 'Copy Button
        SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_COPY, 0, 0
      Case 8 'Paste Button
        cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
        SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_PASTE, 0, 0
        cUndo.sAddText = Clipboard.GetText(vbCFText)
        mdiMain.ActiveForm.AddToUndoStack cUndo
      Case 9 'Delete
        cUndo.lStart = ActiveForm.txtMain.SelStart
        cUndo.sDelText = ActiveForm.txtMain.SelText
        ActiveForm.AddToUndoStack cUndo
        ActiveForm.txtMain.SelText = ""
      Case 14 'URL
        frmInsertURL.Show , Me
      Case 15 'Image
        frmInsertImage.Show , Me
      Case 17 'Left Align
        ActiveForm.InsertTag "<p align=\""left\"" >", "</p>"
      Case 18 'Center Align
        ActiveForm.InsertTag "<p align=\""center\"" >", "</p>"
      Case 19 'Right Align
        ActiveForm.InsertTag "<p align=\""right\"" >", "</p>"
      Case 21 'Comment Code
        ActiveForm.CommentBlock
      Case 22 'Uncomment Code
        ActiveForm.UncommentBlock
      Case 24 'Font
        frmFont.Show , Me
      Case 25 'Bold
        ActiveForm.InsertTag "<B>", "</B>"
      Case 26 'Italic
        ActiveForm.InsertTag "<I>", "</I>"
      Case 27 'Underline
        ActiveForm.InsertTag "<U>", "</U>"
      Case 29
        frmFind.Show , Me
      Case 31
        pictoolbar.Align = 1
      Case 32
        pictoolbar.Align = 2
      Case 33
 
    End Select
End Sub

Private Sub cboRedo_DropDown()
  Dim i As Long
    cboRedo.Clear
    For i = ActiveForm.RedoStack.Count To 1 Step -1
        cboRedo.AddItem "Redo " & ActiveForm.GetUndoText(ActiveForm.RedoStack(i).ModifyType)
    Next
End Sub

Private Sub cboUndo_Click()
  Dim i As Long
    For i = ActiveForm.UndoStack.Count To (ActiveForm.UndoStack.Count - cboUndo.ListIndex) Step -1
        Call cmdUndo_Click
    Next
End Sub

Private Sub cboUndo_DropDown()
  Dim i As Long
    cboUndo.Clear
    For i = ActiveForm.UndoStack.Count To 1 Step -1
        cboUndo.AddItem "Undo " & ActiveForm.GetUndoText(ActiveForm.UndoStack(i).ModifyType)
    Next
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

Private Sub mnu_developers_Click()
OpenIt "http://elucidsoftware.hypermart.net/developer.htm"
End Sub

Private Sub mnu_documents_Click(Index As Integer)
    If WindowCount > 9 Then
       MsgBox "You Must Close Some Windows First"
      Else
         NewDocument
    ActiveForm.txtChanged = -2
    ActiveForm.txtMain.LoadFile mnu_documents(Index).Tag, rtfText
    ActiveForm.txtChanged = 0
    ActiveForm.Text1.Text = mnu_documents(Index).Tag
    ActiveForm.Caption = FileGetName(mnu_documents(Index).Tag)
    End If

End Sub

Private Sub mnu_effectiveperl_Click()
OpenIt "http://www.effectiveperl.com/"
End Sub

Private Sub mnu_exit_Click()
Unload Me
End Sub

Private Sub mnu_homepage_Click()
OpenIt "http://elucidsoftware.hypermart.net"
End Sub

Private Sub mnu_leftwin_Click()
      If leftwin.Visible = True Then
        leftwin.Visible = False
        mnu_leftwin.Checked = False
      Else
        leftwin.Visible = True
        mnu_leftwin.Checked = True
    End If
    
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", mnu_leftwin.Checked
    Set CReg = Nothing

End Sub

Private Sub mnu_mattsriptarchive_Click()
OpenIt "http://www.worldwidemart.com/scripts/"
End Sub

Private Sub mnu_motherofperl_Click()
OpenIt "http://www.webreference.com/perl/"
End Sub

Private Sub mnu_otherservices_Click()
OpenIt "http://elucidsoftware.hypermart.net/services.htm"
End Sub

Private Sub mnu_perlmongers_Click()
OpenIt "http://www.perl.org/"
End Sub

Private Sub mnu_perlreferencelink_Click()
OpenIt "http://www.perl.com/reference/query.cgi?cgi"
End Sub





Private Sub mnu_scriptlocker_Click()
OpenIt "http://www.scriptlocker.com/"
End Sub

Private Sub mnu_settings_Click()
    frmSettings.Show 1, Me
End Sub

Private Sub mnu_software_Click()
OpenIt "http://elucidsoftware.hypermart.net/software.htm"
End Sub

Private Sub mnu_statusbar_Click()

    If picStatusBar.Visible = True Then
        picStatusBar.Visible = False
        mnu_statusbar.Checked = False
      Else
        picStatusBar.Visible = True
        mnu_statusbar.Checked = True
    End If
    
    
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", mnu_statusbar.Checked
    Set CReg = Nothing
End Sub

Private Sub mnu_thecgicollection_Click()
OpenIt "http://www.itm.com/cgicollection/"
End Sub

Private Sub mnu_theperljournal_Click()
OpenIt "http://www.itknowledge.com/tpj/"
End Sub

Private Sub mnu_toolbar_Click()
  Dim CReg As New CRegister
    Set CReg = New CRegister

    If pictoolbar.Visible = True Then

        pictoolbar.Visible = False
        mnu_toolbar.Checked = False
        mnu_hidetoolbar.Checked = False
        If ActiveForm Is Nothing Then
        Else
        ActiveForm.mnu_toolbar.Checked = False
        End If
      Else
        pictoolbar.Visible = True
        mnu_toolbar.Checked = True
        mnu_hidetoolbar.Checked = True
        If ActiveForm Is Nothing Then
        Else
        ActiveForm.mnu_toolbar.Checked = True
        End If
    End If

    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", mnu_toolbar.Checked
    Set CReg = Nothing
End Sub
Private Sub cboRedo_Click()
  Dim i As Long
    For i = ActiveForm.RedoStack.Count To (ActiveForm.RedoStack.Count - cboRedo.ListIndex) Step -1
        Call cmdRedo_Click
    Next
End Sub

Public Sub cmdRedo_Click()
  Dim cRedo As clsUndo
    ActiveForm.bRedoing = False
    If ActiveForm.RedoStack.Count = 0 Then Exit Sub
    '// get the current Redo item
    Set cRedo = ActiveForm.RedoStack(ActiveForm.RedoStack.Count)
    '// add it to the undo stack, and remove it from the Redo stack
    ActiveForm.UndoStack.Add cRedo
    ActiveForm.RedoStack.Remove (ActiveForm.RedoStack.Count)
    '// freeze updates
    LockWindowUpdate ActiveForm.txtMain.hwnd
    '// Redo the text edit
    ActiveForm.txtMain.SelStart = cRedo.lStart
    '// delete any text that was deleted
    ActiveForm.txtMain.SelLength = Len(cRedo.sDelText)
    '// replace the text that was added
    ActiveForm.txtMain.SelText = cRedo.sAddText
    LockWindowUpdate 0
    ActiveForm.bRedoing = True
    ActiveForm.txtMain.SetFocus
End Sub

Public Sub cmdUndo_Click()
  Dim cUndo As clsUndo
    
    If ActiveForm.UndoStack.Count = 0 Then Exit Sub
    '// get the current Undo item
    Set cUndo = ActiveForm.UndoStack(ActiveForm.UndoStack.Count)
    '// add it to the redo stack, and remove it from the Undo stack
    ActiveForm.RedoStack.Add cUndo
    ActiveForm.UndoStack.Remove (ActiveForm.UndoStack.Count)
    '// freeze updates
    LockWindowUpdate ActiveForm.txtMain.hwnd
    '// Undo the text edit
    ActiveForm.txtMain.SelStart = cUndo.lStart
    '// delete any text that was added
    ActiveForm.txtMain.SelLength = Len(cUndo.sAddText)
    '// replace the text that was deleted
    ActiveForm.txtMain.SelText = cUndo.sDelText

    LockWindowUpdate 0
    ' ActiveForm.txtmain.SetFocus

End Sub

Private Sub MDIForm_Resize()
    pictoolbar.Move 0, 0, Me.Width, 400
Toolbar1.Width = Screen.Width
End Sub

Private Sub mnu_File_Click()
    If Not ActiveForm Is Nothing Then
        '       mnu_rename.Enabled = True
        frmMain.mnuFileClose = True
        frmMain.mnuFilePrintSetup.Enabled = True
        frmMain.mnuFilePrint.Enabled = True
      Else
        Exit Sub
    End If

    If ActiveForm.txtChanged = -1 Then
        If Len(ActiveForm.txtMain.FileName) > 0 Then
            frmMain.mnuFileSave.Enabled = True
            frmMain.mnu_revert.Enabled = True
        
            '            mnu_rename.Enabled = True
          Else
            frmMain.mnuFileSave.Enabled = False
            frmMain.mnu_revert.Enabled = False
          
            frmMain.mnuFileSave.Enabled = False
            frmMain.mnu_revert.Enabled = False
            '            mnu_rename.Enabled = False
        End If
        
        frmMain.mnuFileSave.Enabled = True
        frmMain.mnuFileSaveAs.Enabled = True
      Else
        frmMain.mnuFileSave.Enabled = False
        frmMain.mnuFileSaveAs.Enabled = False

    End If
    
    If WindowCount > 1 Then
        'mnu_SaveAll.Enabled = True
      Else
        'mnu_SaveAll.Enabled = False
    End If

  Dim CReg As New CRegister
    Set CReg = New CRegister

End Sub

Private Sub MDIForm_Load()
    LoadSettings
    LoadWindowSettings
    pictoolbar.Move 0, 0, Me.Width, 400
    Picture3.Height = pictoolbar.Height
    Frame1.Left = 0
    Frame1.Width = Screen.Width
    Frame2.Width = Screen.Width
    Frame3.Width = Screen.Width
    Line1.X2 = Screen.Width
    txtStatus.Width = picStatusBar.Width
    'MsgBox "This is a preview release and is only for testing purposes.  We need you to submit any bugs you can find and features you would like to see.  As of now we are aware that some of the features are inoperable.  This is becuase we felt it would be better to get your opinion on how to implement them.", vbInformation, "Preview Release"
    GetRecentList
    GetKeywords
    ShowTips
    TagWindow Me.hwnd
    ParseCommand Command
    Toolbar1.PlaySounds = False
End Sub
Public Sub GotoLine(LineNumber As Long, Highlight As Boolean)
    ActiveForm.GotoLine LineNumber, Highlight
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveWindowSettings
End Sub

Private Sub mnu_New_Click()
    NewDocument
End Sub
Public Sub NewDocument()
  Dim Index As Byte
      If WindowCount > 9 Then
       MsgBox "You Must Close Some Windows First"
      Else
    
    Index = UBound(Document)
    Set Document(Index) = New frmMain
    Document(Index).txtChanged = tsFalse
    Document(Index).Show
    ClearRecentList
    GetRecentList
    End If
End Sub
Private Sub mnu_Open_Click()
       If WindowCount > 9 Then
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
    If ActiveForm Is Nothing Then
        NewDocument
      ElseIf ActiveForm.txtChanged = True Or Len(ActiveForm.txtMain.Text) > 0 Then
        NewDocument
    End If
    ActiveForm.txtChanged = tsnone
    ActiveForm.txtMain.LoadFile CmdDlg.FileName, rtfText
    ActiveForm.Text1.Text = CmdDlg.FileName
    Colorize ActiveForm.txtMain, &H8000&, &HFF0000, &H80&, True
    ActiveForm.Caption = FileGetName(CmdDlg.FileName)
    ActiveForm.txtChanged = tsFalse
    AddRecentList CmdDlg.FileName
    Set CmdDlg = Nothing

End If
End Sub

Private Sub mnu_tools_Click()
OpenIt "http://elucidsoftware.hypermart.net/tools.htm"
End Sub

Private Sub mnu_View_Click()

    If pictoolbar.Visible = True Then
        mnu_toolbar.Checked = True
      Else
        mnu_toolbar.Checked = False
    End If

    If picStatusBar.Visible = True Then
        mnu_statusbar.Checked = True
      Else
        mnu_statusbar.Checked = False
    End If

    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.picLines.Visible = True Then
        frmMain.mnu_linenumbers.Checked = True
      Else
        frmMain.mnu_linenumbers.Checked = False
    End If
End Sub
Public Sub ClearRecentList()
  Dim i  As Integer
  Dim a As Integer
    For i = 1 To mnu_documents.UBound
        Unload mnu_documents(i)
    Next i
    For a = 1 To ActiveForm.mnu_documents.UBound
        Unload ActiveForm.mnu_documents(a)
    Next a
End Sub
Public Sub AddRecentList(FileToAdd As String)
  Dim a As Integer
  Dim sBuf As Variant

  Dim b As Integer
    a = FreeFile

    If FileCheck(App.Path & "\recent.lst") = True Then

        If FileLen(App.Path & "\recent.lst") > 0 Then
            Open App.Path & "\recent.lst" For Input As #a
            sBuf = Input(LOF(a), #a)
            Close #a
        End If

    End If

    b = FreeFile

    Open App.Path & "\recent.lst" For Output As #b
    sBuf = FileToAdd & vbCrLf & sBuf
    Print #b, sBuf
    Close b

    ClearRecentList
    GetRecentList
End Sub
Public Sub GetRecentList()
  Dim a As Integer
  Dim sBuf As String
  Dim NewIndex As Integer
  Dim Count As Integer
  Dim Count1 As Integer
    a = FreeFile
    If FileCheck(App.Path & "\recent.lst") = False Then Exit Sub
    Open App.Path & "\recent.lst" For Input As #a

  Dim i As Integer
    For i = 1 To Me.varDocuments

        If EOF(a) Or i > Me.varDocuments Then GoTo closeit:
        If i > 0 Then mnu_documents(0).Visible = False
        Line Input #a, sBuf
        If FileCheck(sBuf) = False Then GoTo skipit:
        NewIndex = mnu_documents.UBound + 1
        'NewIndex = ActiveForm.mnu_documents.UBound + 1
        Load ActiveForm.mnu_documents(NewIndex)
        Load mnu_documents(NewIndex)
        mnu_documents(NewIndex).Tag = sBuf 'keep the entire path in the tag in case its trimmed
        ActiveForm.mnu_documents(NewIndex).Tag = sBuf 'keep the entire path in the tag in case its trimmed

        If Len(sBuf) > 35 Then
            sBuf = "..." & Right(sBuf, 32)  'make sure this thing isnt to long
        End If
        Count = Count + 1
        Count1 = Count1 + 1
        mnu_documents(NewIndex).Caption = "&" & Count & " " & sBuf
        mnu_documents(NewIndex).Enabled = True
        mnu_documents(NewIndex).Visible = True
        ActiveForm.mnu_documents(NewIndex).Caption = "&" & Count1 & " " & sBuf
        ActiveForm.mnu_documents(NewIndex).Enabled = True
        ActiveForm.mnu_documents(NewIndex).Visible = True
skipit:
    Next i
closeit:
    Close a

End Sub
Public Sub InsertFont(Face As String, Size As Integer, Style As String, lColor As String)
    ActiveForm.InsertFont Face, Size, Style, lColor
End Sub
Sub GetFonts()
Dim i
On Error Resume Next
Combo1.Clear 'Reset list

Combo1.Text = "Times New Roman"
For i = 1 To Screen.FontCount - 1
    Combo1.AddItem Screen.Fonts(i)
    
Next i
End Sub
Public Sub LoadSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister
    varUndoLimit = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UndoLimit", 100)
    varClearUndo = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ClearUndoSave", vbUnchecked)
    varDocuments = Int(CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Documents", 4))
    varUseRelative = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", vbChecked)
    varDefaultFolder = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "DefaultFolder", App.Path)
    pictoolbar.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", True)
    picStatusBar.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True)
    varShowLines = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True)
    leftwin.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", True)
    mnu_leftwin.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", True)
    mnu_toolbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", True)
    mnu_statusbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True)
    frmMain.mnu_linenumbers.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True)
    mnu_formstoolbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowFormsToolbar", False)
   
    Picture3.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowFormsToolbar", False)
    Set CReg = Nothing
End Sub
Public Sub LineNumbers(IsVisible As Boolean)
    ActiveForm.picLines.Visible = IsVisible
    ActiveForm.Form_Resize
End Sub

Private Sub mnu_wwwperlcom_Click()
OpenIt "http://www.perl.com/pub"
End Sub


Private Sub GetKeywords()
  Dim Num%, Buf$
    Num = FreeFile
    Keywords = "|"
    If FileCheck(App.Path & "\keywords.dat") Then
        Open App.Path & "\keywords.dat" For Input As #Num
        While Not EOF(Num)
            Line Input #Num, Buf$
            
            lstKeys.AddItem Left(Buf$, InStr(Buf$, "#") - 1)
            lstKeyDescription.AddItem Mid(Buf$, InStr(Buf$, "#") + 1, Len(Buf$) - InStr(Buf$, "#") + 1)
        Wend
        Close #Num
    End If

End Sub

Private Sub Toolbar1_ButtonMouseOver(ByVal ButtonIndex As Integer, ByVal ButtonKey As String)
txtStatus.Caption = ButtonKey
End Sub
Private Sub ParseCommand(ByVal sCmd As String)
Dim sFile As String
Dim iFIle As Long
Dim sText As String

   If (Me.Visible = False) Then
      Me.Visible = True
   End If

   RestoreAndActivate Me.hwnd

   sCmd = Trim$(sCmd)

   If Len(sCmd) > 0 Then
      If (Left$(sCmd, 1) = """") Then
         sCmd = Mid$(sCmd, 2)
      End If
      If (Right$(sCmd, 1) = """") Then
         sCmd = Left$(sCmd, Len(sCmd) - 1)
      End If
      On Error GoTo ErrorHandler
      sFile = Dir(sCmd, vbNormal)
      iFIle = FreeFile
      Open sCmd For Binary Access Read As #iFIle
      sText = Space$(LOF(iFIle))
      Get #iFIle, , sText
      Close #iFIle
      iFIle = 0

      frmMain.txtMain.Text = sText
    ActiveForm.txtChanged = tsnone
    ActiveForm.Text1.Text = sFile
    Colorize ActiveForm.txtMain, &H8000&, &HFF0000, &H80&, True
    ActiveForm.Caption = FileGetName(sFile)
    ActiveForm.txtChanged = tsFalse
    AddRecentList sFile
   End If
   Exit Sub
   
ErrorHandler:
Dim sErr As String
   sErr = Err.Description
   If (iFIle <> 0) Then
      Close #iFIle
   End If
End Sub

Private Sub Toolbar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     If Button = 2 Then

  Me.PopupMenu Me.mnu_toolbat_popup
 End If
End Sub
Public Sub ShowTips()
    Dim ShowAtStartup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    If Combo1.ListCount = 0 Then
    GetFonts
    End If
    If ShowAtStartup = 1 Then
    frmTip.Show vbModal, Me
    End If
End Sub
