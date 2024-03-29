VERSION 5.00
Begin VB.Form frmInsertURL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert URL"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5505
   Icon            =   "frmInsertURL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5505
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2895
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
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
      ItemData        =   "frmInsertURL.frx":014A
      Left            =   135
      List            =   "frmInsertURL.frx":015D
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
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
      Height          =   375
      Left            =   3855
      TabIndex        =   4
      Top             =   735
      Width           =   1575
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
      Left            =   1815
      TabIndex        =   3
      Top             =   735
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4215
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
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
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label2 
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
      TabIndex        =   5
      Top             =   825
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmInsertURL"
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
'#  Last Modified Date: 10/24/00                                                       #
'#  Webpage: http://elucidsoftware.hypermart.net                                       #
'#                                                                                     #
'#  CgEye - CGI IDE PRODUCTION TOOL                                                    #
'#######################################################################################

Option Explicit
Private FrameType As String

Private Sub chkRelative_Click()

  Dim CReg As New CRegister

    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", chkRelative.Value
    mdiMain.varUseRelative = chkRelative.Value
    Set CReg = Nothing

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

Private Sub Command1_Click()

    If Combo1.ListIndex > 0 Then
        mdiMain.ActiveForm.InsertURL Text1, True, FrameType
      Else
        mdiMain.ActiveForm.InsertURL Text1
    End If
    Unload Me

End Sub

Private Sub Command2_Click()

    Unload Me
    mdiMain.SetFocus

End Sub

Private Sub Command3_Click()

  Dim CmdDlg As New cCommonDialog

    Set CmdDlg = New cCommonDialog
    CmdDlg.Filter = "HTML File(*.htm *.html *.shtml)|*.htm;*.html;*.shtml|Active Server Page(*.asp)|*.asp|All Files(*.*)|*.*"
    CmdDlg.DialogTitle = "Select Image"
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.ShowOpen

    If CmdDlg.FileName = "" Then Exit Sub
    Text1.Text = returnRelPath(mdiMain.varDefaultFolder, CmdDlg.FileName)

End Sub

Private Sub Form_Load()

    Combo1.ListIndex = 0

    If mdiMain.varUseRelative = True Then
        chkRelative.Value = Checked
    End If
    FrameType = ""

End Sub

':) Ulli's Code Formatter V2.0 (11/17/2000 2:46:51 PM) 15 + 74 = 89 Lines
