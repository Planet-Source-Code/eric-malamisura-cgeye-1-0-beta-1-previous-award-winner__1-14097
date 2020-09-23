VERSION 5.00
Begin VB.Form frmTables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Tables"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Make table responce write"
      Height          =   270
      Left            =   90
      TabIndex        =   21
      Top             =   3450
      Visible         =   0   'False
      Width           =   2280
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Insert"
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Layout"
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   5415
      Begin VB.VScrollBar VScroll4 
         Height          =   225
         Left            =   3720
         Max             =   0
         Min             =   -100
         TabIndex        =   25
         Top             =   750
         Width           =   225
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   225
         Left            =   2130
         Max             =   0
         Min             =   -100
         TabIndex        =   24
         Top             =   1470
         Width           =   240
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   225
         Left            =   2130
         Max             =   0
         Min             =   -100
         TabIndex        =   23
         Top             =   1110
         Width           =   240
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   225
         Left            =   2130
         Max             =   0
         Min             =   -100
         TabIndex        =   22
         Top             =   750
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In percent"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   18
         Top             =   915
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In pixels"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   17
         Top             =   645
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   3000
         TabIndex        =   16
         Text            =   "100"
         Top             =   720
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Space Width"
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Text            =   "1"
         Top             =   1440
         Width           =   1200
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1200
         TabIndex        =   9
         Text            =   "1"
         Top             =   1080
         Width           =   1200
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1200
         TabIndex        =   7
         Text            =   "center"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Spacing"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cell padding"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Border size"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Size"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5415
      Begin VB.VScrollBar VScroll6 
         Height          =   240
         Left            =   4755
         Max             =   0
         Min             =   -10000
         TabIndex        =   27
         Top             =   390
         Width           =   255
      End
      Begin VB.VScrollBar VScroll5 
         Height          =   225
         Left            =   2100
         Max             =   0
         Min             =   -10000
         TabIndex        =   26
         Top             =   390
         Width           =   240
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Text            =   "1"
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1185
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Columns"
         Height          =   255
         Left            =   3000
         TabIndex        =   5
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rows"
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   600
      End
   End
   Begin VB.Label lblTable 
      Height          =   255
      Left            =   450
      TabIndex        =   0
      Top             =   6900
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************
' Programmed by Paul B email pbtools@ntlworld.com
'*********************************************************************

Option Explicit ':( Line inserted
Function AddTableAsp(ColumnCount As Long, RowCount As Long, swth As String) As String

    On Error Resume Next

    Dim tmp

    Dim j As Long
    Dim K As Long
    Dim quote$
      quote$ = Chr$(34)
      tmp = "print" & Chr$(34) & "<TABLE Align=\" & Chr$(34) & Combo1.Text & "\" & Chr$(34) & " Border=\" & Chr$(34) & Text3.Text & "\" & Chr$(34) & " cellpadding=\" & Chr$(34) & Text4.Text & "\" & Chr$(34) & swth & " cellspacing=\" & Chr$(34) & Text5.Text & "\" & Chr$(34) & ">\n" & Chr$(34) & ";" & vbCrLf
      For j = 1 To RowCount
          tmp = tmp & vbCrLf & "Print" & Chr$(34) & "<TR>\n" & Chr$(34) & ";" & vbCrLf & "Print" & Chr$(34) & "<TD>Your Text Here</TD>\n" & Chr$(34) & ";" & vbCrLf
          If ColumnCount > 1 Then
              For K = 2 To ColumnCount
                  tmp = tmp & "Print" & Chr$(34) & "<TD>You Text Here</TD>\n" & Chr$(34) & ";" & vbCrLf
              Next K
          End If
          tmp = tmp & "print" & Chr$(34) & "</TR>\n" & Chr$(34) & ";"
    
      Next j

      tmp = tmp & vbCrLf & "print" & Chr$(34) & "</TABLE>\n" & Chr$(34) & ";" & vbCrLf

    Dim cUndo As New clsUndo
      Set cUndo = New clsUndo

      cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
      cUndo.sAddText = tmp
      mdiMain.ActiveForm.AddToUndoStack cUndo
      mdiMain.ActiveForm.txtMain.SelText = tmp

End Function

Private Sub Command1_Click()

  Dim wsizeasp As String
  Dim wsize As String

    wsize = ""
    If Check1.Value = False Then
      Else
        If Option1(0).Value = True Then
            wsize = " width=" & Text6.Text
            wsizeasp = " width=\" & Chr$(34) & Text6.Text & "\" & Chr$(34)
          Else
            wsize = " width=" & Text6.Text & "%"
            wsizeasp = " width=\" & Chr$(34) & Text6.Text & "%" & "\" & Chr$(34)
        End If

    End If
    If Check2.Value = False Then
        Call AddTableAsp(Text1.Text, Text2.Text, wsizeasp)
      Else
        Call AddTableAsp(Text1.Text, Text2.Text, wsizeasp)
    End If
    Unload Me
 
End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    On Error Resume Next

      Combo1.AddItem "Left"
      Combo1.AddItem "Center"
      Combo1.AddItem "Right"

End Sub

Private Sub VScroll1_Change()

  Dim v%

    v = VScroll1.Value * -1
    Text3.Text = Format$(v, "0")

End Sub

Private Sub VScroll2_Change()

  Dim v%

    v = VScroll2.Value * -1
    Text4.Text = Format$(v, "0")

End Sub

Private Sub VScroll3_Change()

  Dim v%

    v = VScroll3.Value * -1
    Text5.Text = Format$(v, "0")

End Sub

Private Sub VScroll4_Change()

  Dim v%

    v = VScroll4.Value * -1
    Text6.Text = Format$(v, "0")

End Sub

Private Sub VScroll5_Change()

  Dim v%

    v = VScroll5.Value * -1
    Text1.Text = Format$(v, "0")

End Sub

Private Sub VScroll6_Change()

  Dim v%

    v = VScroll6.Value * -1
    Text2.Text = Format$(v, "0")

End Sub

':) Ulli's Code Formatter V2.0 (11/17/2000 2:46:37 PM) 3 + 133 = 136 Lines
