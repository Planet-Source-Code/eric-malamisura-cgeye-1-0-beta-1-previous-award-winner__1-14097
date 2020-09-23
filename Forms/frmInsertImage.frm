VERSION 5.00
Begin VB.Form frmInsertImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Image"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "frmInsertImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   6330
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VScroll1 
      Height          =   225
      Left            =   3240
      Max             =   0
      Min             =   -32767
      TabIndex        =   22
      Top             =   1125
      Width           =   240
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   510
      Left            =   3570
      TabIndex        =   21
      Top             =   1230
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   18
      Top             =   1980
      Width           =   6120
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   2805
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Select File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5025
      TabIndex        =   16
      Top             =   2355
      Width           =   1230
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmInsertImage.frx":014A
      Left            =   75
      List            =   "frmInsertImage.frx":015D
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   2685
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      Top             =   2790
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Make in to link"
      Height          =   210
      Left            =   2055
      TabIndex        =   13
      Top             =   1515
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3615
      TabIndex        =   12
      Top             =   1350
      Width           =   1215
   End
   Begin VB.CheckBox chkRelative 
      Alignment       =   1  'Right Justify
      Caption         =   "Get Relative Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   750
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Select File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5040
      TabIndex        =   10
      Top             =   705
      Width           =   1215
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   1425
      Width           =   855
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   1050
      Width           =   855
   End
   Begin VB.CheckBox chkSize 
      Caption         =   "Specify Image Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtBorder 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2955
      TabIndex        =   4
      Text            =   "0"
      Top             =   1095
      Width           =   555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   6165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4935
      TabIndex        =   0
      Top             =   1350
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   20
      Top             =   1740
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Target Frame:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2340
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1425
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   345
      TabIndex        =   6
      Top             =   1065
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Border Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2055
      TabIndex        =   3
      Top             =   1095
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Image File:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmInsertImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'#######################################################################################
'#                            Production Of Elucid Software Inc.                       #
'#                                                                                     #
'#  This has been a production of ES Inc. and is not to be reproduced in any way       #
'#  without permission from ES itself and PB.                                                 #
'#                                                                                     #
'#  Programmer: Eric Malamisura                                                        #
'#  Programmer: Paul Beviss                                                        #
'#  Last Modified Date: 11/9/00                                                       #
'#  Webpage: http://elucidsoftware.hypermart.net                                       #
'#                                                                                     #
'#  CgEye - CGI IDE PRODUCTION TOOL                                                    #
'#######################################################################################
Option Explicit
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private m_Picture As Picture
Private m_bm As BITMAP
Private FrameType As String
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" ( _
                ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Function ImageReadOK(FileName As String) As Boolean

    On Error Resume Next
      Set m_Picture = LoadPicture(FileName)
      If Err Then
          ImageReadOK = False
          Exit Function
      End If
      ImageReadOK = (GetObjectAPI(m_Picture.Handle, Len(m_bm), m_bm) = Len(m_bm))

End Function

Public Property Get WidthPixels() As Long

    WidthPixels = m_bm.bmWidth

End Property

Public Property Get HeightPixels() As Long

    HeightPixels = m_bm.bmHeight

End Property

Public Property Get WidthHiMetric() As Long

    WidthHiMetric = m_Picture.Width

End Property

Public Property Get HeightHiMetric() As Long

    HeightHiMetric = m_Picture.Height

End Property

Private Sub Check1_Click()

    If Me.Height = "2130" Then
        Me.Height = "3600"
        Frame1.Visible = True
      Else
        Me.Height = "2130"
        Frame1.Visible = False
    End If

End Sub

Private Sub chkRelative_Click()

  Dim CReg As New CRegister

    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", chkRelative.Value
    mdiMain.varUseRelative = chkRelative.Value
    Set CReg = Nothing

End Sub

Private Sub chkSize_Click()

    If chkSize.Value = Checked Then
        txtHeight.Enabled = True
        txtWidth.Enabled = True
        txtHeight.BackColor = &H80000005
        txtWidth.BackColor = &H80000005
      Else
        txtWidth.BackColor = &H8000000F
        txtHeight.BackColor = &H8000000F
        txtHeight.Enabled = False
        txtWidth.Enabled = False
    End If

End Sub

Private Sub Command1_Click()

    If chkSize.Value = Checked Then
        mdiMain.ActiveForm.InsertImage Text1, txtBorder, True, txtWidth, txtHeight
      Else
        mdiMain.ActiveForm.InsertImage Text1, txtBorder, False
    End If
    Unload Me
    mdiMain.SetFocus

End Sub

Private Sub Command2_Click()

    Unload Me
    mdiMain.SetFocus

End Sub

Private Sub Command3_Click()

  Dim CmdDlg As New cCommonDialog

    Set CmdDlg = New cCommonDialog

    CmdDlg.Filter = "All Images(*.jpg *.gif)|*.jpg;*.gif|Jpeg(*.jpg)|*.jpg|Gif(*.gif)|*.gif|All Files(*.*)|*.*"
    CmdDlg.DialogTitle = "Select Image"
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.hwnd = mdiMain.hwnd
    CmdDlg.ShowOpen
    If CmdDlg.FileName = "" Then Exit Sub
    ImageReadOK (CmdDlg.FileName)
    txtWidth.Text = WidthPixels
    txtHeight.Text = HeightPixels
  Dim RelativePath As String
  Dim Temp As String
  Dim i As Integer

    Text1.Text = returnRelPath(mdiMain.varDefaultFolder, CmdDlg.FileName)

End Sub

Private Sub Command4_Click()

    Unload Me

End Sub

Private Sub Command5_Click()

  Dim CmdDlg As New cCommonDialog

    Set CmdDlg = New cCommonDialog
    CmdDlg.Filter = "HTML File(*.htm *.html *.shtml)|*.htm;*.html;*.shtml|Active Server Page(*.asp)|*.asp|All Files(*.*)|*.*"
    CmdDlg.DialogTitle = "Select Image"
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.ShowOpen

    If CmdDlg.FileName = "" Then Exit Sub
    Text2.Text = returnRelPath(mdiMain.varDefaultFolder, CmdDlg.FileName)

End Sub

Private Sub Combo1_Click()

    Select Case Combo1.ListIndex
      Case 0
        FrameType = ""
      Case 1
        FrameType = "_self"
      Case 2
        FrameType = "_top"
      Case 3
        FrameType = "_blank"
      Case 4
        FrameType = "_parent"
    End Select

End Sub

Private Sub Command6_Click()

    If Combo1.ListIndex > 0 Then
        mdiMain.ActiveForm.InsertURL Text2, True, FrameType
      Else
        mdiMain.ActiveForm.InsertURL Text2
    End If
    If chkSize.Value = Checked Then
        mdiMain.ActiveForm.InsertImagelink Text1, txtBorder, True, txtWidth, txtHeight
      Else
        mdiMain.ActiveForm.InsertImagelink Text1, txtBorder, False
    End If
    Unload Me
    mdiMain.SetFocus

End Sub

Private Sub Form_Load()

    Combo1.ListIndex = 0
    FrameType = ""
    If mdiMain.varUseRelative = True Then
        chkRelative.Value = Checked
    End If
    Me.Height = "2130"
    SetNumber txtHeight, True
    SetNumber txtWidth, True
    SetNumber txtBorder, True

End Sub

Private Sub VScroll1_Change()

  Dim v%

    v = VScroll1.Value * -1
    txtBorder.Text = Format$(v, "0")

End Sub

':) Ulli's Code Formatter V2.0 (11/17/2000 2:46:53 PM) 28 + 197 = 225 Lines
