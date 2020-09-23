VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9660
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame ctSyntax 
      Caption         =   "Syntax Tooltips"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Left            =   120
      TabIndex        =   31
      Top             =   480
      Width           =   9420
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3480
         TabIndex        =   70
         Text            =   "Combo2"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Edit Keyword Fields"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   65
         Top             =   1200
         Width           =   9135
         Begin VB.ComboBox lstKeywords 
            Height          =   315
            Left            =   120
            TabIndex        =   69
            Text            =   "Select Keyword"
            Top             =   480
            Width           =   2535
         End
         Begin RichTextLib.RichTextBox txtDescription 
            Height          =   735
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   1296
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   1
            TextRTF         =   $"frmSettings.frx":06EA
         End
         Begin VB.Label Label25 
            Caption         =   "Caution: Editing the Keyword Fields can cause problems if not performed correctly.  Should only be used by experienced users only!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2880
            TabIndex        =   71
            Top             =   360
            Width           =   6135
         End
         Begin VB.Label Label14 
            Caption         =   "Keywords:"
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
            TabIndex        =   68
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Tooltip Text:"
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
            TabIndex        =   66
            Top             =   960
            Width           =   1695
         End
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Italic"
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
         Left            =   4440
         TabIndex        =   37
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
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
         Left            =   3480
         TabIndex        =   36
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         ScaleHeight     =   225
         ScaleWidth      =   1065
         TabIndex        =   35
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Text Color:"
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
         Left            =   240
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Font:"
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
         TabIndex        =   33
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame ctHotkeys 
      Caption         =   "Hotkeys"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   9420
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
         ItemData        =   "frmSettings.frx":07E9
         Left            =   3000
         List            =   "frmSettings.frx":07F3
         TabIndex        =   21
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text4 
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
         Left            =   3000
         TabIndex        =   20
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Assign &Key"
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
         Left            =   4320
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         ItemData        =   "frmSettings.frx":0813
         Left            =   5520
         List            =   "frmSettings.frx":0815
         TabIndex        =   18
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2640
         Width           =   9135
      End
      Begin VB.Label Label24 
         Caption         =   "Example: ctrl + a + b =  print ""hotkey assignment set"""
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label Label23 
         Caption         =   $"frmSettings.frx":0817
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   63
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label22 
         Caption         =   "Instructions:"
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
         TabIndex        =   62
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Key Mask"
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
         Left            =   3000
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Trigger Key:"
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
         Left            =   3000
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Assigned Keys"
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
         Left            =   5520
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Text To Assign Key To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   1605
      End
   End
   Begin VB.Frame ctAppearance 
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3450
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   9420
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   255
         Left            =   1080
         TabIndex        =   55
         Top             =   2880
         Width           =   8175
         Begin VB.OptionButton Option2 
            Caption         =   "Never Show"
            Height          =   255
            Left            =   5280
            TabIndex        =   57
            Top             =   0
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Always Show"
            Height          =   255
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   1365
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   300
         Index           =   4
         Left            =   1080
         TabIndex        =   50
         Top             =   2280
         Width           =   8250
         Begin VB.OptionButton chkLinenumbers2 
            Caption         =   "Never Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   54
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton chkLinenumbers1 
            Caption         =   "Always Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   300
         Index           =   3
         Left            =   1080
         TabIndex        =   49
         Top             =   1680
         Width           =   8250
         Begin VB.OptionButton chkStatusbar1 
            Caption         =   "Always Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   1455
         End
         Begin VB.OptionButton chkStatusbar2 
            Caption         =   "Never Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   52
            Top             =   0
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   300
         Index           =   1
         Left            =   1080
         TabIndex        =   46
         Top             =   1080
         Width           =   8250
         Begin VB.OptionButton Option3 
            Caption         =   "Always Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Never Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   51
            Top             =   0
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   300
         Index           =   0
         Left            =   1050
         TabIndex        =   45
         Top             =   465
         Width           =   8250
         Begin VB.OptionButton chkToolbar2 
            Caption         =   "Never Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5280
            TabIndex        =   48
            Top             =   45
            Width           =   1695
         End
         Begin VB.OptionButton chkToolbar1 
            Caption         =   "Always Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   45
            TabIndex        =   47
            Top             =   15
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Label Label20 
         Caption         =   "Toolbar2:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Left Menu"
         Height          =   180
         Left            =   240
         TabIndex        =   38
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label Label8 
         Caption         =   "Status Bar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Line Numbers:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Toolbar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame ctGeneral 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9450
      Begin VB.TextBox txtDefaultFolder 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   8175
      End
      Begin VB.TextBox txtUndo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1680
         TabIndex        =   27
         Text            =   "100"
         Top             =   1680
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Height          =   25
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   9135
      End
      Begin VB.TextBox txtDocuments 
         Alignment       =   2  'Center
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
         Left            =   1440
         TabIndex        =   8
         Text            =   "4"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox chkClearUndo 
         Caption         =   "Clear undo/redo buffer on save"
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
         Left            =   720
         TabIndex        =   7
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox chkRelative 
         Caption         =   "Always use relative paths"
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
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Select"
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
         Left            =   7800
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "Note: Using more undo's/redo's uses up more memory."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   61
         Top             =   2040
         Width           =   3975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Recent Documents:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Undo/Redo Settings:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "time(s) before clearing old undo's/redo's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   28
         Top             =   1680
         Width           =   2880
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Undo/Redo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   26
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label3 
         Caption         =   "recent documents in file menu"
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
         Left            =   1920
         TabIndex        =   11
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Allow up to"
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
         Left            =   600
         TabIndex        =   10
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Root Folder:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame ctheader 
      Caption         =   "Custom header"
      Height          =   3465
      Left            =   120
      TabIndex        =   39
      Top             =   480
      Width           =   9435
      Begin RichTextLib.RichTextBox Custheader 
         Height          =   1365
         Left            =   90
         TabIndex        =   40
         Top             =   225
         Width           =   9240
         _ExtentX        =   16298
         _ExtentY        =   2408
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmSettings.frx":0905
      End
      Begin RichTextLib.RichTextBox Custheader2 
         Height          =   1410
         Left            =   120
         TabIndex        =   41
         Top             =   1830
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   2487
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmSettings.frx":0A10
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Custom header 2"
         Height          =   240
         Left            =   150
         TabIndex        =   42
         Top             =   1635
         Width           =   1815
      End
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
      Left            =   6600
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
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
      Left            =   8040
      TabIndex        =   1
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox syntaxtip 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2850
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   32
      Text            =   "frmSettings.frx":0B1B
      Top             =   1185
      Visible         =   0   'False
      Width           =   3375
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4095
      Left            =   0
      TabIndex        =   60
      Top             =   -15
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7223
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Appearance"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Syntax"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Custom Header"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Custom Inserts"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Hotkeys"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame CtInserts 
      Caption         =   "Custom Inserts"
      Height          =   3420
      Left            =   120
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   9420
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmSettings"
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

Private Sub Custheader_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then

        Me.PopupMenu Me.mnuEdit
    End If

End Sub

Private Sub mnuEditCopy_Click()

    SendMessageLong Custheader.hwnd, WM_COPY, 0, 0

End Sub

Private Sub mnuEditCut_Click()

    SendMessageLong Custheader.hwnd, WM_CUT, 0, 0

End Sub

Private Sub mnuEditDSelectAll_Click()

    With Custheader
        .SetFocus
        .SelStart = 0
        .SelLength = Len(Custheader.Text)
    End With

End Sub

Private Sub mnuEditPaste_Click()

    SendMessageLong Custheader.hwnd, WM_PASTE, 0, 0

End Sub

Public Sub SaveSettings()

  Dim CReg As New CRegister

    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UndoLimit", txtUndo
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ClearUndoSave", chkClearUndo.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Documents", Int(txtDocuments.Text)
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", chkRelative.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "DefaultFolder", txtDefaultFolder.Text
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", chkToolbar1.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", chkStatusbar1.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", chkLinenumbers1.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", Option1.Value
    Set CReg = Nothing

End Sub

Public Sub LoadSettings()

  Dim CReg As New CRegister

    Set CReg = New CRegister
    txtUndo = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UndoLimit", 100)
    chkClearUndo.Value = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ClearUndoSave", vbUnchecked)
    txtDocuments.Text = Int(CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Documents", 4))
    chkRelative.Value = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", vbChecked)
    txtDefaultFolder.Text = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "DefaultFolder", App.Path)

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", True) = True Then
        chkToolbar1.Value = True
      Else
        chkToolbar2.Value = True
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True) = True Then
        chkStatusbar1.Value = True
      Else
        chkStatusbar2.Value = True
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True) = True Then
        chkLinenumbers1.Value = True
      Else
        chkLinenumbers2.Value = True
    End If
    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", True) = True Then
        Option1.Value = True
      Else
        Option2.Value = True
    End If
    Set CReg = Nothing

End Sub

Private Sub Check3_Click()

End Sub

Private Sub Command1_Click()

    SaveSettings

    If Not mdiMain.varDocuments = Int(txtDocuments.Text) Then
        mdiMain.varDocuments = Int(txtDocuments.Text)
        mdiMain.ClearRecentList
        mdiMain.GetRecentList
    End If
    mdiMain.varDefaultFolder = txtDefaultFolder.Text
    
    mdiMain.varClearUndo = chkClearUndo.Value
    mdiMain.varUseRelative = chkRelative.Value
    If FileCheck(App.Path & "\Data\customheader2.txt") Then
        Custheader2.SaveFile App.Path & "\Data\customheader2.txt"
    End If
    If FileCheck(App.Path & "\Data\customheader.txt") Then
        Custheader.SaveFile App.Path & "\Data\customheader.txt"
    End If
    Unload Me
    mdiMain.SetFocus

End Sub

Private Sub Command2_Click()

    Unload Me

End Sub

Private Sub Command5_Click()

  Dim YESNO As String

    YESNO = MsgBox("WARNING: This may cause corruption and data loss or fatal errors by editing this file.  It is not recommended you make changes to this file and it is only for professional use only.  Do you want to continue?", vbCritical + vbYesNo, "Critical File Should Not Be Edited!")
    If YESNO = vbYes Then
        '        UpdateDataList
    End If

End Sub

Private Sub Command3_Click()

    frmFolder.Show 1, Me

End Sub

Private Sub Command6_Click()

    syntaxtip.Visible = True
    syntaxtip.ZOrder 0

End Sub

Private Sub Form_Load()

    LoadSettings
    SetNumber txtUndo, True
    SetNumber txtDocuments, True
    ctGeneral.ZOrder 0
    If FileCheck(App.Path & "\Data\customheader.txt") Then
        Custheader.LoadFile App.Path & "\Data\customheader.txt"
    End If
    If FileCheck(App.Path & "\Data\customheader2.txt") Then
        Custheader2.LoadFile App.Path & "\Data\customheader2.txt"
    End If
  Dim i As Integer
    For i = 0 To mdiMain.lstKeys.ListCount
        lstKeywords.AddItem mdiMain.lstKeys.List(i)
    Next i

End Sub

Private Sub List1_Click()

'    Select Case List1.ListIndex
'      Case 0  ' General
'        ctGeneral.ZOrder 0
'        'ctGeneral.Visible = True
'      Case 1 'Appearance
'        ctAppearance.ZOrder 0
'        'ctAppearance.Visible = True
'      Case 2 'Syntax
'        ctSyntax.ZOrder 0
'      Case 3
'        ctheader.ZOrder 0
'      Case 4
'        CustInserts.ZOrder 0
'    End Select

End Sub

Private Sub lstKeywords_Click()

    txtDescription = mdiMain.lstKeyDescription.List(lstKeywords.ListIndex)

End Sub

Private Sub syntaxtip_Click()

    syntaxtip.Visible = False

End Sub

Private Sub syntaxtip_LostFocus()

    syntaxtip.Visible = False

End Sub

':) Ulli's Code Formatter V2.0 (11/17/2000 2:46:43 PM) 15 + 212 = 227 Lines
Private Sub TabStrip1_Click()
ctGeneral.Visible = False
ctAppearance.Visible = False
ctHotkeys.Visible = False
ctheader.Visible = False
ctSyntax.Visible = False
CtInserts.Visible = False
Select Case TabStrip1.SelectedItem.Index
Case 1 ' General
ctGeneral.Visible = True
Case 2 ' Appearance
ctAppearance.Visible = True
Case 3 ' Syntax
ctSyntax.Visible = True
Case 4 ' Custom Header
ctheader.Visible = True
Case 5 ' Custom Inserts
CtInserts.Visible = True
Case 6 ' Hotkeys
ctHotkeys.Visible = True
End Select
End Sub
