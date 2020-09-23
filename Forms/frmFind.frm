VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboFindWhat 
      Height          =   315
      Left            =   1020
      TabIndex        =   7
      Top             =   180
      Width           =   3735
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Cancel          =   -1  'True
      Caption         =   "&Find"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1500
      Width           =   1335
   End
   Begin VB.ComboBox cboReplaceWith 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   540
      Width           =   3735
   End
   Begin VB.CheckBox chkMatch 
      Caption         =   "Matc&h Case"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   2115
   End
   Begin VB.CheckBox chkWhole 
      Caption         =   "Find &Whole Words Only"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   2115
   End
   Begin VB.Label lblFindWhat 
      Caption         =   "Fi&nd what:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace wi&th:"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   600
      Width           =   1275
   End
End
Attribute VB_Name = "frmFind"
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
'#  Programmer: Paul Beviss                                                       #
'#  Last Modified Date: 11/9/00                                                       #
'#  Webpage: http://elucidsoftware.hypermart.net                                       #
'#                                                                                     #
'#  CgEye - CGI IDE PRODUCTION TOOL                                                    #
'#######################################################################################
Option Explicit ':( Line inserted
Private Const rtfMatchCase = 4
Private Const rtfWholeWord = 2
Dim intPos As Integer

Private Sub cboReplaceWith_Change()

    If Len(cboReplaceWith.Text) > 0 Then
        cmdReplaceAll.Enabled = True
        cmdReplace.Enabled = True
      Else
        cmdReplaceAll.Enabled = False
        cmdReplace.Enabled = False
    End If

End Sub

Private Sub cmdFind_Click()

    If chkMatch.Value = Checked Then
        intPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, , , rtfMatchCase)
    End If

    If chkWhole.Value = Checked Then
        intPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, , , rtfWholeWord)
    End If

    If chkMatch.Value = Unchecked And chkWhole.Value = Unchecked Then
        intPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text)
    End If

    If intPos > 0 Then
        cmdFindNext.Enabled = True
      Else
        cmdFindNext.Enabled = False
    End If

    mdiMain.ActiveForm.txtMain.SetFocus

End Sub

Private Sub cmdFindNext_Click()

    If chkMatch.Value = Checked Then
        intPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, intPos + 1, , rtfMatchCase)
    End If

    If chkWhole.Value = Checked Then
        intPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, intPos + 1, , rtfWholeWord)
    End If

    If chkMatch.Value = Unchecked And chkWhole.Value = Unchecked Then
        intPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, intPos + 1)
    End If

    mdiMain.ActiveForm.txtMain.SetFocus

End Sub

Private Sub cmdReplaceAll_Click()

  Dim tempPos As Integer
  Dim intTimes As Integer
  Dim cUndo As New clsUndo

    Do Until tempPos = -1
        Set cUndo = New clsUndo
        If chkMatch.Value = Checked Then
            tempPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, tempPos + 1, , rtfMatchCase)
        End If

        If chkWhole.Value = Checked Then
            tempPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, tempPos + 1, , rtfWholeWord)
        End If

        If chkMatch.Value = Unchecked And chkWhole.Value = Unchecked Then
            tempPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, tempPos + 1)
        End If

        If tempPos > -1 Then
            cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
            cUndo.sAddText = cboReplaceWith.Text
            mdiMain.ActiveForm.txtMain.SelText = cboReplaceWith.Text
            cUndo.sDelText = cboFindWhat.Text
            intTimes = intTimes + 1
            cUndo.ModifyType = ReplaceText
            mdiMain.ActiveForm.AddToUndoStack cUndo
        End If

    Loop

    MsgBox "Replaced " & intTimes & " item(s)", vbInformation, "Replace"

End Sub

':) Ulli's Code Formatter V2.0 (11/17/2000 2:47:00 PM) 16 + 92 = 108 Lines
