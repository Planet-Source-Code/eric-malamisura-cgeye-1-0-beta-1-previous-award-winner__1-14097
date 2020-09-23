VERSION 5.00
Begin VB.Form frminsertforms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Forms"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Panel6 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2535
      Left            =   0
      TabIndex        =   53
      Top             =   0
      Visible         =   0   'False
      Width           =   4665
      Begin VB.CommandButton Command6 
         Caption         =   "&Insert"
         Height          =   375
         Left            =   3330
         TabIndex        =   62
         Top             =   1830
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   61
         Top             =   1830
         Width           =   975
      End
      Begin VB.Frame Frame6 
         Caption         =   "Form"
         Height          =   1680
         Left            =   60
         TabIndex        =   54
         Top             =   45
         Width           =   4245
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   810
            TabIndex        =   57
            Top             =   345
            Width           =   3270
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   795
            TabIndex        =   56
            Top             =   1140
            Width           =   3300
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frminsertforms.frx":0000
            Left            =   810
            List            =   "frminsertforms.frx":000A
            TabIndex        =   55
            Text            =   "post"
            Top             =   720
            Width           =   3300
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Action"
            Height          =   255
            Left            =   135
            TabIndex        =   60
            Top             =   375
            Width           =   495
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "enctype"
            Height          =   270
            Left            =   135
            TabIndex        =   59
            Top             =   1140
            Width           =   630
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Method"
            Height          =   285
            Left            =   135
            TabIndex        =   58
            Top             =   750
            Width           =   660
         End
      End
   End
   Begin VB.Frame Panel5 
      BorderStyle     =   0  'None
      Caption         =   "Frame9"
      Height          =   2505
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   4650
      Begin VB.CommandButton Command7 
         Caption         =   "&Insert"
         Height          =   375
         Left            =   3315
         TabIndex        =   52
         Top             =   1905
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   5
         Left            =   2280
         TabIndex        =   51
         Top             =   1905
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "Check Box"
         Height          =   1695
         Left            =   105
         TabIndex        =   44
         Top             =   105
         Width           =   4215
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   1095
            TabIndex        =   48
            Top             =   360
            Width           =   2775
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   1080
            TabIndex        =   47
            Top             =   840
            Width           =   2775
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Not Checked"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   46
            Top             =   1230
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Checked"
            Height          =   255
            Index           =   3
            Left            =   2520
            TabIndex        =   45
            Top             =   1230
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   735
         End
      End
   End
   Begin VB.Frame Panel4 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2520
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   4650
      Begin VB.CommandButton Command4 
         Caption         =   "&Insert"
         Height          =   360
         Left            =   3300
         TabIndex        =   42
         Top             =   1905
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   41
         Top             =   1905
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Push button"
         Height          =   1695
         Left            =   90
         TabIndex        =   33
         Top             =   75
         Width           =   4215
         Begin VB.OptionButton Option3 
            Caption         =   "Reset"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   38
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Submit"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   37
            Top             =   1200
            Width           =   975
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   36
            Top             =   1200
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1320
            TabIndex        =   35
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   1320
            TabIndex        =   34
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Value / Label"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.Frame Panel3 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2505
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton Command3 
         Caption         =   "&Insert"
         Height          =   375
         Left            =   3390
         TabIndex        =   31
         Top             =   1905
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   1
         Left            =   2325
         TabIndex        =   30
         Top             =   1905
         Width           =   975
      End
      Begin VB.Frame Frame3 
         Caption         =   "Option Button"
         Height          =   1695
         Left            =   135
         TabIndex        =   23
         Top             =   75
         Width           =   4215
         Begin VB.OptionButton Option2 
            Caption         =   "Selected"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   27
            Top             =   1200
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Not Selected"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   26
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   1080
            TabIndex        =   25
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame Panel2 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2400
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4560
      Begin VB.CommandButton Command2 
         Caption         =   "&Insert"
         Height          =   375
         Left            =   3330
         TabIndex        =   21
         Top             =   1830
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   2
         Left            =   2265
         TabIndex        =   20
         Top             =   1830
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   "Scolling text box"
         Height          =   1695
         Left            =   90
         TabIndex        =   11
         Top             =   45
         Width           =   4215
         Begin VB.VScrollBar VScroll2 
            Height          =   225
            Left            =   3675
            TabIndex        =   70
            Top             =   1335
            Width           =   255
         End
         Begin VB.VScrollBar VScroll1 
            Height          =   225
            Left            =   1500
            TabIndex        =   69
            Top             =   1335
            Width           =   225
         End
         Begin VB.TextBox Text6 
            Height          =   495
            Left            =   960
            MultiLine       =   -1  'True
            TabIndex        =   15
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   960
            TabIndex        =   14
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   960
            TabIndex        =   13
            Text            =   "0"
            Top             =   1305
            Width           =   795
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   3120
            TabIndex        =   12
            Text            =   "0"
            Top             =   1305
            Width           =   840
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Number of lines"
            Height          =   255
            Left            =   1920
            TabIndex        =   19
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Width"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   615
         End
      End
   End
   Begin VB.Frame Panel1 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4500
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   3
         Left            =   2265
         TabIndex        =   9
         Top             =   1845
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Insert"
         Height          =   375
         Left            =   3315
         TabIndex        =   8
         Top             =   1845
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Text Input 1 line"
         Height          =   1695
         Left            =   105
         TabIndex        =   1
         Top             =   60
         Width           =   4215
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1080
            TabIndex        =   7
            Top             =   435
            Width           =   3015
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            Top             =   960
            Width           =   3015
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Password input"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   3
            Top             =   1320
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Text Input"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   2
            Top             =   1320
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Value"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Text Name"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   975
         End
      End
   End
   Begin VB.Frame Panel7 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2685
      Left            =   15
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   4635
      Begin VB.CommandButton Command8 
         Caption         =   "Insert"
         Height          =   375
         Left            =   3360
         TabIndex        =   68
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Cancel"
         Height          =   375
         Index           =   6
         Left            =   2310
         TabIndex        =   67
         Top             =   1860
         Width           =   975
      End
      Begin VB.Frame Frame7 
         Caption         =   "Label"
         Height          =   1695
         Left            =   135
         TabIndex        =   64
         Top             =   60
         Width           =   4215
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   3855
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Text"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "frminsertforms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************
' Programmed by Paul B email pbtools@ntlworld.com
'*********************************************************************

Private Sub Command5_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command6_Click()
Dim Form
Dim encode
If Text14.Text = "" Then
Else
encode = " enctype=\""" & Text14.Text & "\"""
End If
Form = "<form method=\""" & Combo1.Text & "\"" action=\""" & Text11.Text & "\""" & encode & " >"
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sAddText = Form
    mdiMain.ActiveForm.AddToUndoStack cUndo
mdiMain.ActiveForm.txtMain.SelText = Form
Unload Me
End Sub

Private Sub Command7_Click()

If Option2(2).Value = False Then
inputtext = "<input type=\""checkbox\"" name=\""" & Text12.Text & "\"" value=\""" & Text13.Text & "\"" checked>"
Else
inputtext = "<input type=\""checkbox\"" name=\""" & Text12.Text & "\"" value=\""" & Text13.Text & "\"">"
End If
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sAddText = inputtext
    mdiMain.ActiveForm.AddToUndoStack cUndo
mdiMain.ActiveForm.txtMain.SelText = inputtext
Unload Me
End Sub

Private Sub Command8_Click()
Dim Label
Label = "<Label>" & Text15.Text & "</Label>"
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sAddText = Label
    mdiMain.ActiveForm.AddToUndoStack cUndo
mdiMain.ActiveForm.txtMain.SelText = Label
Unload Me
End Sub


Private Sub Command1_Click()

If Option1(1).Value = True Then
inputtext = "<Input type=\""Password\"" name=\""" & Text1.Text & "\"" Value=\""" & Text2.Text & "\"">"
Else
inputtext = "<Input type=\""Text\"" name=\""" & Text1.Text & "\"" Value=\""" & Text2.Text & "\"">"
End If
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sAddText = inputtext
    mdiMain.ActiveForm.AddToUndoStack cUndo
mdiMain.ActiveForm.txtMain.SelText = inputtext

Unload Me
End Sub

Private Sub Command2_Click()

inputtext = "<p><textarea rows=\""" & Text5.Text & "\"" name=\""" & Text3.Text & "\"" cols=\""" & Text4.Text & "\"">" & Text6.Text & "</textarea></p>"
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sAddText = inputtext
    mdiMain.ActiveForm.AddToUndoStack cUndo
mdiMain.ActiveForm.txtMain.SelText = inputtext
Unload Me
End Sub

Private Sub Command3_Click()
If Option2(1).Value = True Then
inputtext = "<input type=\""radio\"" name=\""" & Text7.Text & "\"" value=\""" & Text8.Text & "\"" checked>"
Else
inputtext = "<input type=\""radio"" name=\""" & Text7.Text & "\"" value=\""" & Text8.Text & "\"">"
End If
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sAddText = inputtext
    mdiMain.ActiveForm.AddToUndoStack cUndo
mdiMain.ActiveForm.txtMain.SelText = inputtext

Unload Me
End Sub

Private Sub Command4_Click()

 If Option3(0).Value = True Then
 inputtext = "<Input type=\""Button\"" name=\""" & Text9.Text & "\"" Value=\""" & Text10.Text & "\"">"
 Else
  If Option3(1).Value = True Then
  inputtext = "<Input type=\""submit\"" name=\""" & Text9.Text & "\"" Value=\""" & Text10.Text & "\"">"
  Else
  inputtext = "<Input type=\""reset\"" name=\""" & Text9.Text & "\"" Value=\""" & Text10.Text & "\"">"
End If
End If
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
    cUndo.sAddText = inputtext
    mdiMain.ActiveForm.AddToUndoStack cUndo
mdiMain.ActiveForm.txtMain.SelText = inputtext

Unload Me
End Sub

Private Sub Form_Load()
Me.Move ScaleHeight / 2, ScaleWidth / 2
End Sub

Private Sub VScroll1_Change()
    Dim V%
    V = VScroll1.Value
    Text4.Text = Format(V, "0")
End Sub

Private Sub VScroll2_Change()
    Dim V%

    V = VScroll2.Value
    Text5.Text = Format(V, "0")
End Sub
