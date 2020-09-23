VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H008A8A8A&
   Caption         =   "CgEye By Elucid Software"
   ClientHeight    =   7545
   ClientLeft      =   180
   ClientTop       =   735
   ClientWidth     =   11355
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0472
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture5 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6450
      Left            =   11340
      ScaleHeight     =   6450
      ScaleWidth      =   15
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   15
      Begin VB.PictureBox Picture6 
         Height          =   315
         Left            =   -30
         ScaleHeight     =   255
         ScaleWidth      =   450
         TabIndex        =   26
         Top             =   0
         Width           =   510
         Begin VB.CommandButton Command1 
            Enabled         =   0   'False
            Height          =   270
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox leftwin 
      Align           =   3  'Align Left
      Height          =   6450
      Left            =   0
      ScaleHeight     =   6390
      ScaleWidth      =   2235
      TabIndex        =   6
      Top             =   840
      Width           =   2295
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   5565
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   2145
         Begin MSComctlLib.Toolbar Toolbar5 
            Height          =   330
            Left            =   1200
            TabIndex        =   42
            Top             =   120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            ButtonWidth     =   609
            Style           =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Root Directory"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "New Folder"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "&Open"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   5160
            Width           =   975
         End
         Begin VB.FileListBox lstFile 
            Height          =   1650
            Left            =   120
            Pattern         =   "*.cgi;*.pl"
            TabIndex        =   40
            Top             =   3360
            Width           =   2055
         End
         Begin VB.DirListBox lstDir 
            Height          =   2790
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   2055
         End
         Begin VB.DriveListBox lstDrive 
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   0
         TabIndex        =   29
         Top             =   350
         Visible         =   0   'False
         Width           =   2175
         Begin VB.ComboBox Combo3 
            Height          =   4470
            ItemData        =   "mdiMain.frx":05DA
            Left            =   240
            List            =   "mdiMain.frx":08C0
            Style           =   1  'Simple Combo
            TabIndex        =   30
            Top             =   120
            Width           =   1800
         End
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame5"
         Height          =   5490
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   2205
         Begin VB.ListBox lstBookmark 
            Height          =   645
            Left            =   240
            TabIndex        =   13
            Top             =   2640
            Width           =   1815
         End
         Begin VB.ListBox lstSub 
            Height          =   2010
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Jump To Sub"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   120
            Width           =   1755
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Bookmarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   2370
            Width           =   1755
         End
      End
      Begin VB.ListBox lstSubPos 
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   6240
         Visible         =   0   'False
         Width           =   135
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   688
         ButtonWidth     =   609
         Appearance      =   1
         _Version        =   393216
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   150
            Left            =   45
            ScaleHeight     =   150
            ScaleWidth      =   1725
            TabIndex        =   35
            Top             =   30
            Width           =   1725
            Begin VB.Label lblHeader 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Header"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   360
               TabIndex        =   36
               Top             =   -15
               Width           =   1485
            End
         End
         Begin VB.Frame Frame7 
            Height          =   120
            Left            =   0
            TabIndex        =   15
            Top             =   180
            Width           =   10785
         End
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   6285
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   11086
         Placement       =   1
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   3
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Files"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Navigation"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "&Commands"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lstKeyDescription 
         Height          =   450
         Left            =   1215
         TabIndex        =   34
         Top             =   5910
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.ListBox lstKeys 
         Height          =   645
         Left            =   735
         TabIndex        =   33
         Top             =   5730
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   11355
      TabIndex        =   4
      Top             =   435
      Width           =   11355
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Left            =   3885
         TabIndex        =   31
         Top             =   60
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         ButtonWidth     =   609
         Style           =   1
         ImageList       =   "imlToolbarIcons(0)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "color"
               ImageKey        =   "color"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "web_link"
               ImageKey        =   "web_link2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "picture"
               ImageKey        =   "picture"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar customins 
         Height          =   330
         Left            =   3045
         TabIndex        =   24
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   582
         ButtonWidth     =   609
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imlToolbarIcons(4)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Back"
               Object.ToolTipText     =   "Back"
               ImageKey        =   "cus"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Forward"
               Object.ToolTipText     =   "Forward"
               ImageKey        =   "Forward"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   330
         Left            =   45
         TabIndex        =   16
         Top             =   45
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   582
         ButtonWidth     =   609
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imlToolbarIcons(1)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "RECTANGL"
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "RECTANGL"
               ImageKey        =   "RECTANGL"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "BUTTON"
               ImageKey        =   "BUTTON"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "textinput"
               ImageKey        =   "textinput"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "muilttext"
               ImageKey        =   "muilttext"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "check"
               ImageKey        =   "check"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "optionbutton"
               ImageKey        =   "optionbutton"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "combo"
               ImageKey        =   "combo"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "font"
               ImageKey        =   "font"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   -15
         TabIndex        =   5
         Top             =   -90
         Width           =   11325
      End
      Begin RichTextLib.RichTextBox cust 
         Height          =   45
         Left            =   3825
         TabIndex        =   9
         Top             =   120
         Visible         =   0   'False
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   79
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"mdiMain.frx":1036
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
      ScaleWidth      =   11355
      TabIndex        =   2
      Top             =   0
      Width           =   11355
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   15
         TabIndex        =   17
         Top             =   60
         Width           =   11310
         _ExtentX        =   19950
         _ExtentY        =   582
         ButtonWidth     =   609
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imlToolbarIcons(3)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   26
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "-"
               Style           =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Description     =   "New"
               Object.ToolTipText     =   "New"
               ImageKey        =   "new2"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Description     =   "Open"
               Object.ToolTipText     =   "Open"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Description     =   "Save"
               Object.ToolTipText     =   "Save"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "-"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Description     =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Description     =   "Copy"
               Object.ToolTipText     =   "Copy"
               Object.Tag             =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Description     =   "Paste"
               Object.ToolTipText     =   "Paste"
               Object.Tag             =   "Paste"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Description     =   "Delete"
               Object.ToolTipText     =   "Delete"
               Object.Tag             =   "Delete"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Description     =   "-"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   1200
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   4
               Object.Width           =   2390
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "insfont"
               ImageKey        =   "qfont"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               ImageKey        =   "left"
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "center"
               ImageKey        =   "center"
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "right"
               ImageKey        =   "right"
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "font"
               ImageKey        =   "font"
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "bold"
               ImageKey        =   "b"
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "i"
               ImageKey        =   "i"
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "underline"
               ImageKey        =   "u"
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "find"
               ImageKey        =   "find"
            EndProperty
         EndProperty
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "mdiMain.frx":1135
            Left            =   5910
            List            =   "mdiMain.frx":1151
            TabIndex        =   23
            Text            =   "1"
            Top             =   15
            Width           =   570
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   4095
            Sorted          =   -1  'True
            TabIndex        =   22
            Text            =   "Times New Roman"
            Top             =   15
            Width           =   1785
         End
         Begin VB.CommandButton cmdUndo 
            Height          =   255
            Left            =   2790
            Picture         =   "mdiMain.frx":116D
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Undo Last Event"
            Top             =   45
            Width           =   330
         End
         Begin VB.ComboBox cboUndo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   "Undo"
            Top             =   15
            Width           =   630
         End
         Begin VB.CommandButton cmdRedo 
            Height          =   255
            Left            =   3420
            MaskColor       =   &H000000C0&
            Picture         =   "mdiMain.frx":12B7
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Redo Last Undo"
            Top             =   45
            Width           =   330
         End
         Begin VB.ComboBox cboRedo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3390
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Redo"
            Top             =   15
            Width           =   630
         End
      End
      Begin VB.Frame Frame1 
         Height          =   120
         Left            =   15
         TabIndex        =   3
         Top             =   -90
         Width           =   11295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   30
         X2              =   9450
         Y1              =   420
         Y2              =   420
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11355
      TabIndex        =   0
      Top             =   840
      Width           =   11355
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   11355
      TabIndex        =   1
      Top             =   840
      Width           =   11355
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   1920
      Top             =   6630
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1401
            Key             =   "RECTANGL"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1513
            Key             =   "BUTTON"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1625
            Key             =   "textinput"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1A37
            Key             =   "muilttext"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":20B9
            Key             =   "check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2417
            Key             =   "optionbutton"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":28A1
            Key             =   "combo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2E43
            Key             =   "font"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   3
      Left            =   2505
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":2F55
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3067
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3179
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":328B
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":339D
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":34AF
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":35C1
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":36D3
            Key             =   "color"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3AA7
            Key             =   "link"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3FAB
            Key             =   "picture"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":43DB
            Key             =   "left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":44F3
            Key             =   "center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":460B
            Key             =   "right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4723
            Key             =   "b"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":483B
            Key             =   "i"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4953
            Key             =   "u"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4A6B
            Key             =   "link2"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4BC7
            Key             =   "font"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4D23
            Key             =   "find"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4E7F
            Key             =   "qfont"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4FDF
            Key             =   "new2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   4
      Left            =   3075
      Top             =   6645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":513F
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5251
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5363
            Key             =   "cus"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   0
      Left            =   3660
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":54BF
            Key             =   "color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5891
            Key             =   "web_link"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5A6B
            Key             =   "picture"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5E99
            Key             =   "web_link2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   32
      Top             =   7290
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   9253
            MinWidth        =   7057
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "12/24/00"
         EndProperty
      EndProperty
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
         Caption         =   "&Insert Toolbar"
         Checked         =   -1  'True
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
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_about 
         Caption         =   "&About"
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

Private Sub cmdOpen_Click()
ActiveForm.txtMain.LoadFile lstFile.Path & "\" & lstFile.FileName, rtfText
End Sub

Private Sub lstDir_Change()
 lstFile.Path = lstDir.Path
End Sub

Private Sub lstDrive_Change()
lstDir.Path = lstDrive.Drive
End Sub

Private Sub lstFile_DblClick()
cmdOpen_Click
End Sub

Private Sub lstSub_DblClick()
Dim iStart As Long
Dim iLength As Long
Dim Buff As String
Buff = lstSubPos.List(lstSub.ListIndex)

iStart = Int(Trim(Left(Buff, InStr(Buff, " ")))) + 1
iLength = Int(Trim(Right(Buff, InStr(Buff, " ")))) - 1

mdiMain.ActiveForm.txtMain.SelStart = iStart
mdiMain.ActiveForm.txtMain.SelLength = iLength - iStart
mdiMain.ActiveForm.txtMain.SetFocus
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "color"
         Dim CmdDlg As New cCommonDialog
         Set CmdDlg = New cCommonDialog
         On Error GoTo Cleanup
         CmdDlg.Color = 0
         CmdDlg.CancelError = True
         CmdDlg.ShowColor
         ActiveForm.InsertTag FormatRGBString(CmdDlg.Color), ""
Cleanup:
        Case "web_link"
         frmInsertURL.Show , Me
        Case "picture"
         frmInsertImage.Show , Me
    End Select
End Sub

Private Sub customins_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    Select Case Button.Key
        Case "Back"
        If FileCheck(App.Path & "\Data\customheader.txt") Then
        cust.LoadFile App.Path & "\Data\customheader.txt"
        ActiveForm.InsertTag cust.Text, "", False
        End If
        Case "Forward"
        If FileCheck(App.Path & "\Data\customheader2.txt") Then
        cust.LoadFile App.Path & "\Data\customheader2.txt"
        ActiveForm.InsertTag cust.Text, "", False
        End If
    End Select
End Sub

Private Sub mnu_about_Click()
frmAbout.Show 1, mdiMain
End Sub

Private Sub TabStrip1_Click()

If TabStrip1.Tabs(1).Selected = True Then
Frame6.Visible = True
lblHeader.Caption = "Files"
Else
Frame6.Visible = False
End If


If TabStrip1.Tabs(2).Selected = True Then
lblHeader.Caption = "Navigation"
Frame5.Visible = True
Else
Frame5.Visible = False
End If

If TabStrip1.Tabs(3).Selected = True Then
lblHeader.Caption = "Command List"
Frame4.Visible = True
Else
Frame4.Visible = False
End If

End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
      If Button.Index > 3 Then If ActiveForm Is Nothing Then Exit Sub
    
        Dim cUndo As clsUndo
        Set cUndo = New clsUndo
    
    Select Case Button.Key
      Case "New" 'New Button
       NewDocument
      Case "Open" 'Open Button
        mnu_Open_Click
      Case "Save" 'Save Button
        If ActiveForm.txtChanged = -1 Then
            If Len(ActiveForm.txtMain.FileName) > 0 Then
                mdiMain.ActiveForm.mnuFileSave_Click
              Else
                mdiMain.ActiveForm.mnuFileSaveAs_Click
            End If
        End If
      Case "Cut" 'Cut Button
        cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
        cUndo.sDelText = mdiMain.ActiveForm.txtMain.SelText
        SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_CUT, 0, 0
        cUndo.ModifyType = CutText
        mdiMain.ActiveForm.AddToUndoStack cUndo
      Case "Copy" 'Copy Button
        SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_COPY, 0, 0
      Case "Paste" 'Paste Button
        cUndo.lStart = mdiMain.ActiveForm.txtMain.SelStart
        SendMessageLong mdiMain.ActiveForm.txtMain.hwnd, WM_PASTE, 0, 0
        cUndo.sAddText = Clipboard.GetText(vbCFText)
        cUndo.ModifyType = PasteText
        ActiveForm.AddToUndoStack cUndo
      Case "Delete" 'Delete
        cUndo.lStart = ActiveForm.txtMain.SelStart
        cUndo.sDelText = ActiveForm.txtMain.SelText
        ActiveForm.AddToUndoStack cUndo
        ActiveForm.txtMain.SelText = ""

      Case "Left" 'Left Align
        ActiveForm.InsertTag "<p align=\""left\"" >", "</p>"
      Case "center" 'Center Align
        ActiveForm.InsertTag "<p align=\""center\"" >", "</p>"
      Case "right" 'Right Align
        ActiveForm.InsertTag "<p align=\""right\"" >", "</p>"
      Case "insfont"
        mdiMain.InsertFont Combo1.Text, Combo2.Text, "", ""
      Case "font" 'Font
        frmFont.Show , Me
      Case "bold" 'Bold
        ActiveForm.InsertTag "<B>", "</B>"
      Case "i" 'Italic
        ActiveForm.InsertTag "<I>", "</I>"
      Case "underline" 'Underline
        ActiveForm.InsertTag "<U>", "</U>"
      Case "find"
        frmFind.Show , Me
    End Select
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    Select Case Button.Key
        Case "RECTANGL"
        frminsertforms.Show 0, Me
        frminsertforms.Panel6.Visible = True
        Case "BUTTON"
        frminsertforms.Show 0, Me
        frminsertforms.Panel4.Visible = True
        Case "textinput"
        frminsertforms.Show 0, Me
        frminsertforms.Panel1.Visible = True
        Case "muilttext"
        frminsertforms.Show 0, Me
        frminsertforms.Panel2.Visible = True
        Case "check"
        frminsertforms.Show 0, Me
        frminsertforms.Panel5.Visible = True
        Case "optionbutton"
        frminsertforms.Show 0, Me
        frminsertforms.Panel3.Visible = True
        Case "combo"
        frminsertforms.Show 0, Me
        Case "font"
        frminsertforms.Show 0, Me
        frminsertforms.Panel7.Visible = True
    End Select

End Sub

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







Private Sub Command2_Click()
ActiveForm.CheckSyntax
End Sub

Private Sub Combo3_DblClick()
ActiveForm.InsertTag Combo3.Text & "", ""
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
frmTemplate.Show 0, Me
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

Private Sub pictoolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then

  Me.PopupMenu Me.mnu_toolbat_popup
 End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     If Button = 2 Then

  Me.PopupMenu Me.mnu_toolbat_popup2
 End If
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
    ActiveForm.txtMain.SetFocus
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
    ListSubs
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

   ' If picStatusBar.Visible = True Then
    '    picStatusBar.Visible = False
     '   mnu_statusbar.Checked = False
     ' Else
      '  picStatusBar.Visible = True
      '  mnu_statusbar.Checked = True
   ' End If
    
    
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
    ActiveForm.txtMain.SetFocus
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
    ActiveForm.txtMain.SetFocus

End Sub

Private Sub MDIForm_Resize()
If leftwin.Height <= 60 Then Exit Sub
pictoolbar.Move 0, 0, Me.Width, 400
Toolbar1.Width = Screen.Width
TabStrip1.Move 0, 0, leftwin.Width, leftwin.Height - 60
End Sub
Private Sub leftwin_Resize()
If leftwin.Height / 2 <= 1300 Then Exit Sub 'make sure it dont crash from division by zero

    TabStrip1.Height = leftwin.Height - 100
    'Controls
    Frame4.Top = 300
    Frame5.Top = 300
    Frame6.Top = 300
    Frame4.Height = TabStrip1.Height - 650
    Frame5.Height = Frame4.Height
    Frame6.Height = Frame4.Height
    
    Frame4.Width = TabStrip1.Width
    Frame6.Width = Frame4.Width
    Frame5.Width = Frame4.Width
    Combo3.Height = Frame4.Height - 200
    lstSub.Height = Frame4.Height / 2 - 400
    lstBookmark.Top = lstSub.Height + 600
    lstBookmark.Height = Frame4.Height / 2 - 400
    Label3.Top = lstBookmark.Top - 250
    lstDir.Height = leftwin.Height / 2 - 400
    lstFile.Height = leftwin.Height / 2 - 1300
    lstFile.Top = lstDir.Height + 500
    cmdOpen.Top = lstFile.Top + lstFile.Height + 100
'    txtStatus.Width = picStatusBar.Width
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
  'Preview Expiration Testing Code

    If Format$(Date, "MMDDYY") >= 123100 Then
  Dim YESNO As String
        YESNO = MsgBox("This version of CgEye is a preview release and has expired." & vbCrLf & vbCrLf & "Would you like to visit Elucid Software Webpage to obtain the latest version now?", vbQuestion + vbYesNo, "Expired Version!")
        If YESNO = vbYes Then OpenIt "http://elucidsoftware.hypermart.net"
        End
    End If
    'End of Expiration Testing Code
    
    Me.Caption = "CgEye By Elucid Software"
    LoadSettings
    LoadWindowSettings
    ResizeControls
    GetRecentList
    GetKeywords
    TagWindow Me.hwnd
    ParseCommand Command
    lblHeader.Caption = "Files"
    
    
    Frame1.Left = 0
    Frame1.Width = Screen.Width
'    Frame2.Width = Screen.Width
    Frame3.Width = Screen.Width
    Line1.X2 = Screen.Width
    
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
ListSubs
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

   ' If picStatusBar.Visible = True Then
    '    mnu_statusbar.Checked = True
    '  Else
     '   mnu_statusbar.Checked = False
   ' End If

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

    If FileCheck(App.Path & "\Data\recent.lst") = True Then

        If FileLen(App.Path & "\Data\recent.lst") > 0 Then
            Open App.Path & "\Data\recent.lst" For Input As #a
            sBuf = Input$(LOF(a), #a)
            Close #a
        End If

    End If

    b = FreeFile

    Open App.Path & "\Data\recent.lst" For Output As #b
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
'  Dim Count1 As Integer

    a = FreeFile
    If FileCheck(App.Path & "\Data\recent.lst") = False Then Exit Sub
    Open App.Path & "\Data\recent.lst" For Input As #a

  Dim i As Integer
    For i = 1 To Me.varDocuments

        If EOF(a) Or i > Me.varDocuments Then GoTo closeit:
        
        If i > 0 Then
        mnu_documents(0).Visible = False
        ActiveForm.mnu_documents(0).Visible = False
        End If
        Line Input #a, sBuf
        If FileCheck(sBuf) = False Then GoTo skipit:
        NewIndex = mnu_documents.UBound + 1
        NewIndex = ActiveForm.mnu_documents.UBound + 1
        Load mnu_documents(NewIndex)
        Load ActiveForm.mnu_documents(NewIndex)
        mnu_documents(NewIndex).Tag = sBuf 'keep the entire path in the tag in case its trimmed
        ActiveForm.mnu_documents(NewIndex).Tag = sBuf 'keep the entire path in the tag in case its trimmed

        If Len(sBuf) > 35 Then
            sBuf = "..." & Right$(sBuf, 32)  'make sure this thing isnt to long
        End If
        Count = Count + 1
'        Count1 = Count1 + 1
        mnu_documents(NewIndex).Caption = "&" & Count & " " & sBuf
        mnu_documents(NewIndex).Enabled = True
        mnu_documents(NewIndex).Visible = True
        ActiveForm.mnu_documents(NewIndex).Caption = "&" & Count & " " & sBuf
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
   ' picStatusBar.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True)
    varShowLines = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True)
    leftwin.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", True)
    mnu_leftwin.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Showleftprop", True)
    mnu_toolbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", True)
    mnu_statusbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True)
    frmMain.mnu_linenumbers.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True)
    mnu_formstoolbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowFormsToolbar", False)
   
    Picture3.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowFormsToolbar", False)
    Set CReg = Nothing
    
Dim ShowAtStartup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    If Combo1.ListCount = 0 Then
    GetFonts
    End If
    If ShowAtStartup = 1 Then
    frmTip.Show vbModal, Me
    End If
    
    
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
    If FileCheck(App.Path & "\data\keywords.dat") Then
        Open App.Path & "\data\keywords.dat" For Input As #Num
        While Not EOF(Num)
            Line Input #Num, Buf$
            
            lstKeys.AddItem Left$(Buf$, InStr(Buf$, "#") - 1)
            lstKeyDescription.AddItem Mid$(Buf$, InStr(Buf$, "#") + 1, Len(Buf$) - InStr(Buf$, "#") + 1)
        Wend
        Close #Num
    End If

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

Private Sub ResizeControls()

    pictoolbar.Move 0, 0, Me.Width, 400
    Toolbar1.Width = Screen.Width
    pictoolbar.Move 0, 0, Me.Width, 400

'    Picture3.Height = pictoolbar.Height
'    TabStrip1.Height = leftwin.Height - 300

End Sub

Public Sub ListSubs()
Dim a As Long
Dim b As Long
Dim Buff As String

lstSub.Clear
lstSubPos.Clear
a = 1
LockWindowUpdate ActiveForm.txtMain.hwnd
Do Until a < 1
'intPos = mdiMain.ActiveForm.txtMain.Find(cboFindWhat.Text, intPos + 1, , rtfWholeWord)

a = mdiMain.ActiveForm.txtMain.Find("sub", a + 1, , rtfWholeWord)

If a < 1 Then
    GoTo done:
Else
   b = InStr(a, mdiMain.ActiveForm.txtMain.Text, "{")
    If b < 1 Then
    GoTo skipadd:
    Else
    Buff = Mid(mdiMain.ActiveForm.txtMain.Text, a + 1, b - a - 1)
    End If
End If
lstSub.AddItem Buff
lstSubPos.AddItem a + 3 & " " & b - 1
skipadd:
Loop
done:
ActiveForm.txtMain.SelStart = 1
LockWindowUpdate 0
ActiveForm.txtMain.SetFocus
End Sub

Private Sub Toolbar5_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
lstDir.Path = lstDrive.Drive & "\"
Case 2
Dim a As String
a = InputBox("Enter Directory Name", "Directory Name")
If Not a = "" Then
    MkDir lstDir.Path & "\" & a
    lstDir.Refresh
End If
End Select
End Sub
