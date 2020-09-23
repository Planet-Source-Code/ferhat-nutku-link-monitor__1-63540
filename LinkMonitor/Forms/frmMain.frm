VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13590
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9510
   ScaleWidth      =   13590
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSaveFile 
      Left            =   11280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrRowCount 
      Interval        =   250
      Left            =   12240
      Top             =   0
   End
   Begin MSComctlLib.StatusBar sbarRowCount 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   9255
      Width           =   13590
      _ExtentX        =   23971
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtMonitor 
      Height          =   285
      Left            =   12840
      TabIndex        =   15
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer tmrMonitor 
      Interval        =   500
      Left            =   11760
      Top             =   0
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8775
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   15478
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Link Monitor"
      TabPicture(0)   =   "frmMain.frx":1042
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(3)"
      Tab(0).Control(1)=   "Label1(2)"
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(4)=   "Label1(5)"
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(6)=   "cmbCategory"
      Tab(0).Control(7)=   "txtAddress"
      Tab(0).Control(8)=   "txtInfo"
      Tab(0).Control(9)=   "chkUserPass"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUserID"
      Tab(0).Control(11)=   "txtPassword"
      Tab(0).Control(12)=   "btnAddAddress"
      Tab(0).Control(13)=   "btnAddCategory"
      Tab(0).Control(14)=   "txtNewCategory"
      Tab(0).Control(15)=   "btnMonitor"
      Tab(0).Control(16)=   "Frame1"
      Tab(0).Control(17)=   "Frame2"
      Tab(0).Control(18)=   "btnEditCategory"
      Tab(0).Control(19)=   "btnCancelCategory"
      Tab(0).Control(20)=   "btnDeleteCategory"
      Tab(0).Control(21)=   "pbrEncrypt"
      Tab(0).Control(22)=   "btnDecryptDB"
      Tab(0).Control(23)=   "btnEncryptDB"
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "WEB Surf"
      TabPicture(1)   =   "frmMain.frx":105E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Line3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Line4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Line5"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Line6"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Line7"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line8"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmbBrowseCat"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtBrowseUserName"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtBrowsePassword"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "cmbBrowseTitle"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmbBrowseAddress"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "dtLastVisit"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "fgridAddresses"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "btnAll"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "btnDeleteTitle"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "btnDeleteLink"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "btnSaveLinks"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "My Favorite Links"
      TabPicture(2)   =   "frmMain.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkTitle"
      Tab(2).Control(1)=   "chkAddress"
      Tab(2).Control(2)=   "chkCategory"
      Tab(2).Control(3)=   "chkUserName"
      Tab(2).Control(4)=   "chkPassword"
      Tab(2).Control(5)=   "chkFirstVisitDate"
      Tab(2).Control(6)=   "chkLastVisitDate"
      Tab(2).Control(7)=   "chkTotalVisit"
      Tab(2).Control(8)=   "btnViewLinks"
      Tab(2).Control(9)=   "btnBrowseLinks"
      Tab(2).Control(10)=   "cmbFavoriteLinks"
      Tab(2).Control(11)=   "fgridFavAddresses"
      Tab(2).Control(12)=   "Line2"
      Tab(2).Control(13)=   "Line1"
      Tab(2).Control(14)=   "Label10"
      Tab(2).Control(15)=   "Label9"
      Tab(2).ControlCount=   16
      Begin VB.CheckBox chkTitle 
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   59
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CheckBox chkAddress 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72000
         TabIndex        =   58
         Top             =   5520
         Width           =   1455
      End
      Begin VB.CheckBox chkCategory 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         TabIndex        =   57
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CheckBox chkUserName 
         Caption         =   "UserName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         TabIndex        =   56
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CheckBox chkPassword 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72000
         TabIndex        =   55
         Top             =   6840
         Width           =   1455
      End
      Begin VB.CheckBox chkFirstVisitDate 
         Caption         =   "FirstVisitDate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70320
         TabIndex        =   54
         Top             =   5040
         Width           =   1695
      End
      Begin VB.CheckBox chkLastVisitDate 
         Caption         =   "LastVisitDate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70320
         TabIndex        =   53
         Top             =   5400
         Width           =   1935
      End
      Begin VB.CheckBox chkTotalVisit 
         Caption         =   "TotalVisit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70320
         TabIndex        =   52
         Top             =   5880
         Width           =   1815
      End
      Begin VB.CommandButton btnSaveLinks 
         Caption         =   "Save Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6720
         TabIndex        =   51
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton btnViewLinks 
         Caption         =   "View Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73800
         TabIndex        =   50
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton btnBrowseLinks 
         Caption         =   "Browse Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -73800
         TabIndex        =   49
         Top             =   5520
         Width           =   1455
      End
      Begin VB.ComboBox cmbFavoriteLinks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmMain.frx":1096
         Left            =   -74760
         List            =   "frmMain.frx":1098
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   5040
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgridFavAddresses 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   47
         Top             =   1560
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   4895
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton btnDeleteLink 
         Height          =   255
         Left            =   7920
         Picture         =   "frmMain.frx":109A
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton btnDeleteTitle 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3960
         Picture         =   "frmMain.frx":687C
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton btnEncryptDB 
         Caption         =   "Encrypt Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   44
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CommandButton btnDecryptDB 
         Caption         =   "Decrypt Database"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   43
         Top             =   6120
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar pbrEncrypt 
         Height          =   255
         Left            =   -74640
         TabIndex        =   41
         Top             =   8280
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton btnDeleteCategory 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -62880
         TabIndex        =   34
         Top             =   2340
         Width           =   850
      End
      Begin VB.CommandButton btnCancelCategory 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -63840
         TabIndex        =   33
         Top             =   2340
         Width           =   850
      End
      Begin VB.CommandButton btnEditCategory 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -64800
         TabIndex        =   32
         Top             =   2340
         Width           =   850
      End
      Begin VB.CommandButton btnAll 
         Caption         =   "All Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5520
         TabIndex        =   31
         Top             =   480
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Navigation Browser Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69360
         TabIndex        =   28
         Top             =   4440
         Width           =   2775
         Begin VB.OptionButton rbtnFireFoxNav 
            Caption         =   "Mozilla Firefox"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   2295
         End
         Begin VB.OptionButton rbtnIExplorerNav 
            Caption         =   "Internet Explorer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Monitoring Browser Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -72120
         TabIndex        =   25
         Top             =   4440
         Width           =   2655
         Begin VB.OptionButton rbtnFireFox 
            Caption         =   "Mozilla Firefox"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton rbtnIExplorer 
            Caption         =   "Internet Explorer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1815
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgridAddresses 
         Height          =   6615
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   12840
         _ExtentX        =   22648
         _ExtentY        =   11668
         _Version        =   393216
         FixedCols       =   0
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker dtLastVisit 
         Height          =   285
         Left            =   11800
         TabIndex        =   23
         Top             =   1380
         Width           =   1250
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         Format          =   48955393
         CurrentDate     =   38583
      End
      Begin VB.ComboBox cmbBrowseAddress 
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         TabIndex        =   22
         Top             =   1380
         Width           =   3855
      End
      Begin VB.ComboBox cmbBrowseTitle 
         Height          =   315
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   21
         Top             =   1380
         Width           =   4000
      End
      Begin VB.TextBox txtBrowsePassword 
         Height          =   285
         Left            =   10800
         TabIndex        =   20
         Top             =   1380
         Width           =   900
      End
      Begin VB.TextBox txtBrowseUserName 
         Height          =   285
         Left            =   9800
         TabIndex        =   19
         Top             =   1380
         Width           =   925
      End
      Begin VB.ComboBox cmbBrowseCat 
         Height          =   315
         Left            =   8300
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   1380
         Width           =   1400
      End
      Begin VB.CommandButton btnMonitor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stop Monitoring"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   16
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox txtNewCategory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68880
         TabIndex        =   14
         Top             =   2340
         Width           =   2175
      End
      Begin VB.CommandButton btnAddCategory 
         Caption         =   "Add Category"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66600
         TabIndex        =   13
         Top             =   2340
         Width           =   1700
      End
      Begin VB.CommandButton btnAddAddress 
         Caption         =   "Add Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   3
         Top             =   4560
         Width           =   1815
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72120
         TabIndex        =   12
         Top             =   3900
         Width           =   3255
      End
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72120
         TabIndex        =   11
         Top             =   3420
         Width           =   3255
      End
      Begin VB.CheckBox chkUserPass 
         Caption         =   "User Name && Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74640
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3060
         Width           =   2535
      End
      Begin VB.TextBox txtInfo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72120
         TabIndex        =   0
         Top             =   1380
         Width           =   10000
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72120
         TabIndex        =   1
         Top             =   1860
         Width           =   10000
      End
      Begin VB.ComboBox cmbCategory 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -72120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2340
         Width           =   2895
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000C0&
         X1              =   11760
         X2              =   13080
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000C0&
         X1              =   10800
         X2              =   11640
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000C0&
         X1              =   9840
         X2              =   10680
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000C0&
         X1              =   8280
         X2              =   9600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000C0&
         X1              =   4320
         X2              =   8160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000C0&
         X1              =   240
         X2              =   4200
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         X1              =   -72000
         X2              =   -68880
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         X1              =   -74760
         X2              =   -72360
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Label Label10 
         Caption         =   "Favorite Links"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   -74760
         TabIndex        =   61
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "Save Link Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   -72000
         TabIndex        =   60
         Top             =   4560
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Encrypt / Decrypt Status"
         Height          =   255
         Left            =   -74640
         TabIndex        =   42
         Top             =   7920
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Visit Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   11760
         TabIndex        =   40
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   10800
         TabIndex        =   39
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   9840
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8280
         TabIndex        =   37
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Address of The Web Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Title of The Web Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   -74640
         TabIndex        =   9
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   -74640
         TabIndex        =   8
         Top             =   3480
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Title of The Web Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -74640
         TabIndex        =   7
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Link of The Web Site"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   -74640
         TabIndex        =   6
         Top             =   1920
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   -74640
         TabIndex        =   5
         Top             =   2400
         Width           =   825
      End
   End
   Begin VB.Menu mnGridPopup 
      Caption         =   "Grid Popup"
      Visible         =   0   'False
      Begin VB.Menu mnBrowse 
         Caption         =   "Browse"
      End
      Begin VB.Menu mnEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnSysTrayMenu 
      Caption         =   "SysTrayMenu"
      Visible         =   0   'False
      Begin VB.Menu mnShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'BUGS
'when clicked in the grid update result are not shown immediately
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Enum ComboBoxName
   BrowseTitle = 1
   BrowseAddress = 2
End Enum
Private cmbIndex As ComboBoxName
Private ResizableControls() As ResizableControl



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public goTray As CTray
Attribute goTray.VB_VarHelpID = -1


Public categoryInEdit As Boolean
Public goBrowseSql As String
Public gomonitorBrowserType As String
Public gonavigationBrowserType As String
Public gIsDataEncrypted As String
Public WithEvents goRecordset As Recordset
Attribute goRecordset.VB_VarHelpID = -1
Public goConnection As New Connection
Public goDialog As New CDialog
Public goIni As New CIni
Public goIO As New CIO
Public goAdo As New CADO
Public goControl As New CControl
Public goCLink As New CLink
Public goSecurity As New CSecurity
Public goBlowfish As New CBlowfish
Public goCString As New CString


'Search Keys
Public gtitleKey As String
Public gaddressKey As String
Public gcategoryKey As String
Public guserNameKey As String
Public gpasswordKey As String
'Search Keys

Public gDBName As String
Public ggridMouseRow As Integer
Public ggridMouseColumn As Integer
Public gmouseButton As MouseButtonConstants



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
   Dim sSql As String
    
   'Load settings from INI file
   gomonitorBrowserType = goIni.INIGetSetting(App.Path & "\" & con_INI_File, "Program", "MonitorBrowserType")
   gonavigationBrowserType = goIni.INIGetSetting(App.Path & "\" & con_INI_File, "Program", "NavigationBrowserType")
   gIsDataEncrypted = goIni.INIGetSetting(App.Path & "\" & con_INI_File, "Program", "DataEncrypted")
   
   'Labels
   Label1(4).Visible = False
   Label1(5).Visible = False
   
    
   'Buttons
   btnAddAddress.Default = True
   btnCancelCategory.Enabled = False
   'Encryption buttons
   If gIsDataEncrypted = "1" Then
      btnEncryptDB.Enabled = False
      btnDecryptDB.Enabled = True
   ElseIf gIsDataEncrypted = "0" Then
      btnEncryptDB.Enabled = True
      btnDecryptDB.Enabled = False
   Else
      btnEncryptDB.Enabled = True
      btnDecryptDB.Enabled = True
   End If
     
        
   'ComboBoxes
   For i = 2 To 20 Step 2
      cmbFavoriteLinks.AddItem (CStr(i))
   Next
   cmbFavoriteLinks.ListIndex = 0
      
   Call goIO.LoadComboBox(cmbCategory, MApplication.Path & "\" & "Categories.dat")
   Call goIO.LoadComboBox(cmbBrowseAddress, MApplication.Path & "\" & "BrowsedAddresses.dat")
   Call goIO.LoadComboBox(cmbBrowseTitle, MApplication.Path & "\" & "BrowsedTitles.dat")
      
   
   'CheckBoxes
   chkTitle.Value = 1
   chkAddress.Value = 1
   chkCategory.Value = 1
   
   
   
   'Datetime Pickers
   dtLastVisit.Value = DateTime.Date
                
                
   'Form
   Me.Caption = MApplication.ProductNameVersion
   
   
   'RadioButtons
   'Load browser type for monitoring
   If (gomonitorBrowserType = "Firefox") Then
      rbtnFireFox.Value = True
   ElseIf (gomonitorBrowserType = "IE") Then
      rbtnIExplorer.Value = True
   Else
      rbtnIExplorer.Value = True
   End If
   
   'Load browser type for navigation
   If (gonavigationBrowserType = "Firefox") Then
      rbtnFireFoxNav.Value = True
   ElseIf (gonavigationBrowserType = "IE") Then
      rbtnIExplorerNav.Value = True
   Else
      rbtnIExplorerNav.Value = True
   End If
   
     
   'TextBoxes
   txtMonitor.Visible = False
   txtUserID.Visible = False
   txtPassword.Visible = False
   txtUserID.Text = " "
   txtPassword.Text = " "
   
    
   'StatusBars
   Call goControl.AddPanelsToStatusBar(sbarRowCount, 4)
      
   sbarRowCount.Panels(1).Style = sbrDate
   sbarRowCount.Panels(1).ToolTipText = MResourse.rToday
   
   sbarRowCount.Panels(2).Style = sbrTime
   sbarRowCount.Panels(2).ToolTipText = MResourse.rTime
   
   sbarRowCount.Panels(3).AutoSize = sbrSpring
   sbarRowCount.Panels(4).AutoSize = sbrSpring
   sbarRowCount.Panels(4).Text = rTotal & " " & cmbCategory.ListCount & " " & rCategory
       
       
   'SSTabs
   tabMain.Tab = 0
      
    
   'Database
   gDBName = MApplication.Path & "\" & MResourse.con_DB_Name
       
    
   'MSHFlexGrid
   fgridAddresses.Redraw = True
   fgridAddresses.Sort = flexSortGenericAscending
    
    
   'Set Width of MSHFlexGrid columns
   Call SetMSHFlexColumnWidth(fgridAddresses)
   With fgridFavAddresses
      .ColWidth(0) = 3500
      .ColWidth(1) = 3000
   End With
     
     
     
   'Fill datagrid.
   sSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount  FROM Links ORDER BY Title"
   Call FillDataGrid
    
   'Fill the global goRecordSet
   Set goRecordset = goAdo.GetRecordSetByFileName(gDBName, sSql)
    
   'Set the text of the Header Columns of MSHFlexGrid
   Call SetMSHFlexGridHeaderText(fgridAddresses)
  
   
   'Form Resize Initialization
   ReDim ResizableControls(frmMain.Count) As ResizableControl
   On Local Error Resume Next
   For i = 0 To frmMain.Count - 1
      'ResizableControls(i).left = frmMain.Controls(i).left
      ResizableControls(i).width = frmMain.Controls(i).width
      'ResizableControls(i).top = frmMain.Controls(i).top
      ResizableControls(i).height = frmMain.Controls(i).height
      'ResizableControls(i).fontsize = frmMain.Controls(i).fontsize
   Next
   
   
   'System Tray Icon Properties
   Set goTray = New CTray
   Call goTray.GiveHandle(Me.hWnd)
     
   'Add Icon
   Call goTray.AddTrayIcon(LoadPicture(MApplication.Path & "\" & MResourse.rSysTrayIcon))
   
   
   'Decrypt database
   'DecryptDB
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

        Select Case Button

            Case MouseButtonConstants.vbLeftButton
               Me.Show
               Me.WindowState = FormWindowStateConstants.vbNormal

            Case MouseButtonConstants.vbRightButton
               Call PopupMenu(mnSysTrayMenu)

            Case Else

       End Select

End Sub

Private Sub Form_Resize()

   'System Tray Icon Part
   Select Case WindowState
        
        'Form Minimized
        Case vbMinimized
           Me.Hide
        
        'Form Maximized
        Case vbMaximized
        
        'Form Normal
        Case vbNormal
         
   End Select


   Static originalSizeX, originalSizeY
   Dim i, ratioX, ratioY
   
   If originalSizeX = 0 Then
      originalSizeX = frmMain.width
      originalSizeY = frmMain.height
      Exit Sub
   End If
   
   If WindowState <> 1 Then
      ratioX = frmMain.width / originalSizeX
      ratioY = frmMain.height / originalSizeY
   
      If (ratioX < 1 And ratioY < 1) Then
         Exit Sub
      End If
   
      On Local Error Resume Next
   
      For i = 0 To frmMain.Count - 1
         'Controls(i).left = ResizableControls(i).left * ratioX
         Controls(i).width = ResizableControls(i).width * ratioX
         'Controls(i).top = ResizableControls(i).top * ratioY
         Controls(i).height = ResizableControls(i).height * ratioY
         'Controls(i).fontsize = ResizableControls(i).fontsize * ratioY
      Next
   
   End If
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Save keywords
   Call goIO.SaveComboBox(cmbCategory, MApplication.Path & "\" & "Categories.dat")
   Call goIO.SaveComboBox(cmbBrowseAddress, MApplication.Path & "\" & "BrowsedAddresses.dat")
   Call goIO.SaveComboBox(cmbBrowseTitle, MApplication.Path & "\" & "BrowsedTitles.dat")
   
   'Save browser type for monitoring setting
   If (rbtnFireFox.Value) Then
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "MonitorBrowserType", "Firefox")
   ElseIf (rbtnIExplorer.Value) Then
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "MonitorBrowserType", "IE")
   Else
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "MonitorBrowserType", "IE")
   End If
   
   'Save browser type for navigation setting
   If (rbtnFireFoxNav.Value) Then
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "NavigationBrowserType", "Firefox")
   ElseIf (rbtnIExplorerNav.Value) Then
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "NavigationBrowserType", "IE")
   Else
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "NavigationBrowserType", "IE")
   End If
   
   'Save Encryption status
   If (btnEncryptDB.Enabled And Not btnDecryptDB.Enabled) Then
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "DataEncrypted", "0")
   ElseIf (Not btnEncryptDB.Enabled And btnDecryptDB.Enabled) Then
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "DataEncrypted", "1")
   Else
      Call goIni.INISaveSetting(App.Path & "\" & con_INI_File, "Program", "DataEncrypted", "")
   End If
   
   'Ýkonu Sil
   Call goTray.RemoveTrayIcon
   End
   
End Sub

Private Sub SetMSHFlexGridHeaderText(ByRef pflexgrid As MSHFlexGrid)
   'Sets the text of the Header Columns of MSHFlexGrid
   'Call SetMSHFlexGridHeaderText(fgridAddresses)
    
   Dim sHeaderText(7) As String
   sHeaderText(0) = rTitle
   sHeaderText(1) = rAddress
   sHeaderText(2) = rCategory
   sHeaderText(3) = rUserName
   sHeaderText(4) = rPassword
   sHeaderText(5) = rLastDate
   sHeaderText(6) = rFirstDate
   sHeaderText(7) = rVisitCount
   For i = 0 To UBound(sHeaderText)
        pflexgrid.TextMatrix(0, i) = sHeaderText(i)
   Next
   
   
End Sub

Private Sub SetMSHFlexColumnWidth(ByRef pflexgrid As MSHFlexGrid)
    'Set Width of MSHFlexGrid columns
    
    For i = 0 To fgridAddresses.Cols
      pflexgrid.ColWidth(i) = 4000
    Next
    pflexgrid.ColWidth(2) = 1500
    pflexgrid.ColWidth(5) = 1500
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''
'SSTab 1 Codes
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkUserPass_Click()
    Dim sart As Boolean
    sart = 0
    If chkUserPass = 1 Then sart = 1
    Label1(4).Visible = sart
    Label1(5).Visible = sart
    txtUserID.Visible = sart
    txtPassword.Visible = sart
End Sub

Private Sub mnClose_Click()
   'Ýkonu Sil
   Call goTray.RemoveTrayIcon
   End
End Sub

Private Sub mnShow_Click()
   Me.Show
   Me.WindowState = FormWindowStateConstants.vbNormal
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
   
   'If second tab is selected
   If (tabMain.Tab = 1) Then
      FillcmbBrowseCat
   End If
      
      
   'If third tab is selected
   If (tabMain.Tab = 2) Then
      Call LoadFavoriteLinks(cmbFavoriteLinks.Text)
   End If
End Sub

Private Sub LoadFavoriteLinks(ptop As String)
      Dim sSql As String
      
      'Fill datagrid.
      sSql = "SELECT TOP " & ptop & " Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links ORDER BY VisitCount DESC"
      Call goControl.FillMSHFlexGrid(fgridFavAddresses, gDBName, goRecordset, sSql)
End Sub

Private Sub txtNewCategory_Change()
   If Trim(txtNewCategory.Text) = "" Then
      btnAddCategory.Enabled = False
   Else
      btnAddCategory.Enabled = True
   End If
End Sub

Private Sub btnAddCategory_Click()
    Call goControl.AddItemToComboBoxUnique(cmbCategory, goSecurity.RemoveUnSecureChars(txtNewCategory.Text, True))
    sbarRowCount.Panels(4).Text = rTotal & " " & cmbCategory.ListCount & " " & rCategory
    txtNewCategory = ""
End Sub

Private Sub btnEditCategory_Click()
   Dim categoryName As String
   categoryName = goSecurity.RemoveUnSecureChars(txtNewCategory.Text, True)

   If categoryInEdit Then
   
      'Do not allow empty string
      If (categoryName <> "") Then
         UpdateCategory (categoryName)
      Else
         Exit Sub
      End If
      btnEditCategory.Caption = "Edit"
      btnCancelCategory.Enabled = False
      btnDeleteCategory.Enabled = True
      txtNewCategory.Text = Empty
      
   Else
   
      txtNewCategory.Text = cmbCategory.Text
      txtNewCategory.SetFocus
      btnEditCategory.Caption = "Save"
      btnCancelCategory.Enabled = True
      btnDeleteCategory.Enabled = False
      
   End If
   
   categoryInEdit = Not categoryInEdit
  
End Sub

Private Sub btnCancelCategory_Click()
  
   If categoryInEdit Then
      txtNewCategory.Text = Empty
      btnEditCategory.Caption = "Edit"
      btnCancelCategory.Enabled = False
      btnDeleteCategory.Enabled = True
   End If
   categoryInEdit = Not categoryInEdit
End Sub

Private Sub btnDeleteCategory_Click()
   Dim deleteCategory As String
   deleteCategory = cmbCategory.Text

   rMsg = MsgBox("Do you want to delete " & deleteCategory & " category ?", vbYesNo + vbQuestion, MApplication.ProductNameVersion + " - Delete confirmation")
   If rMsg = vbNo Then
      Exit Sub
   End If
   
   'Delete selected category
   Call goControl.RemoveItemFromComboBox(cmbCategory, deleteCategory)
   
   'Select first item of combobox
   cmbCategory.ListIndex = 1
   
   'Count total number of categories
   sbarRowCount.Panels(4).Text = rTotal & " " & cmbCategory.ListCount & " " & rCategory
End Sub

Private Sub UpdateCategory(ByVal pnewCategoryName)
   Dim oldCategoryName As String
   Dim newCategoryName As String
   oldCategoryName = cmbCategory.Text
   newCategoryName = pnewCategoryName
   
   On Local Error GoTo RollBack:
   
   'Remove the old category name from combobox
   Call goControl.RemoveItemFromComboBox(cmbCategory, oldCategoryName)
   

   'Update the global recordset, increase VisitCount & update LastDate
   sSql = "UPDATE Links SET Category = '" & newCategoryName & "' WHERE Category = '" & oldCategoryName & "';"
   Call goAdo.ExecuteCommand(gDBName, sSql)
   Call goControl.AddItemToComboBoxUnique(cmbCategory, newCategoryName)
   FillcmbBrowseCat
   FillDataGrid
   Exit Sub
   
RollBack:
   Call MsgBox("An error occurred when updating the category.", vbExclamation + vbOKOnly, MApplication.ProductNameVersion + " - Error information")
End Sub

Private Sub btnMonitor_Click()
On Local Error GoTo ExitSub
    
    If btnMonitor.Caption = rStartMonitoring Then
        btnMonitor.Caption = rStopMonitoring
        Call goTray.ChangeTrayIcon(LoadPicture(MApplication.Path & "\" & MResourse.rSysTrayIcon))
    Else
        btnMonitor.Caption = rStartMonitoring
        
        'If txtAddress is full then do not write "http://"
        If (txtAddress.Text = "") Then
            txtAddress.Text = "http://"
        End If
       
        Call goTray.ChangeTrayIcon(LoadPicture(MApplication.Path & "\" & MResourse.rSysTrayDisIcon))
    End If
    
    If tmrMonitor.Enabled = True Then
        tmrMonitor.Enabled = False
    Else
        tmrMonitor.Enabled = True
    End If
    
ExitSub:
End Sub

Private Sub btnAddAddress_Click()
   On Local Error Resume Next
     
   Dim strTitle As String
   Dim strAddress As String
   Dim strUserName As String
   Dim strPassword As String
   Dim strCategory As String
      
      
   strTitle = goSecurity.RemoveUnSecureChars(txtInfo.Text, True)
   strAddress = goSecurity.RemoveUnSecureChars(txtAddress.Text, True)
   strUserName = goSecurity.RemoveUnSecureChars(txtUserID.Text, True)
   strPassword = goSecurity.RemoveUnSecureChars(txtPassword.Text, True)
   strCategory = goSecurity.RemoveUnSecureChars(cmbCategory.Text, True)
     
     
   'Do not add empty data
   If strTitle = "" Or strAddress = "" Then
      Exit Sub
   End If
   
   
   'Do not add same links into db
   bcontrolSame = ControlSameLink(strAddress)
   If bcontrolSame Then
      Call MsgBox("There is already an item exists with a link " & strAddress, vbOKOnly + vbExclamation, MApplication.ProductNameVersion + " - Link Add Warning")
      Exit Sub
   End If
      
     
   goRecordset.AddNew
   goRecordset![Title] = strTitle
   goRecordset![Address] = strAddress
   goRecordset![Category] = strCategory
   goRecordset![userName] = strUserName
   goRecordset![password] = strPassword
   goRecordset![FirstDate] = DateTime.Now()
   goRecordset![VisitCount] = 1
   goRecordset.Update
   goRecordset.Requery 'Requery refreshs the datagrid
       
   'Fill datagrid.
   Set fgridAddresses.DataSource = goRecordset
   
   'Set the text of the Header Columns of MSHFlexGrid
   Call SetMSHFlexGridHeaderText(fgridAddresses)
   
   'Clear textboxes
   txtAddress.Text = ""
   txtInfo.Text = ""
   txtUserID.Text = ""
   txtPassword.Text = ""
End Sub

Private Sub btnEncryptDB_Click()
   EncryptDB
   Call goControl.CreateFlipFlop(btnEncryptDB, btnDecryptDB)
End Sub

Private Sub btnDecryptDB_Click()
   DecryptDB
   Call goControl.CreateFlipFlop(btnDecryptDB, btnEncryptDB)
End Sub

Public Sub EncryptDB()
   Dim oRecordSet As New ADODB.Recordset
   Dim sSql As String
   Dim strFieldValue
   Dim recordCount, i As Double
   i = 1
   
   sSql = "SELECT * FROM Links"
   Call oRecordSet.Open(sSql, goAdo.GetConnectionAccess(gDBName), adOpenStatic, adLockOptimistic)
   
   recordCount = oRecordSet.recordCount

   While oRecordSet.EOF <> True
      
      oRecordSet!Title = goBlowfish.EncryptString(IIf(IsNull(oRecordSet!Title), "", oRecordSet!Title), MResourse.ENCRYPT_KEY)
      oRecordSet!Address = goBlowfish.EncryptString(IIf(IsNull(oRecordSet!Address), "", oRecordSet!Address), MResourse.ENCRYPT_KEY)
      oRecordSet!Category = goBlowfish.EncryptString(IIf(IsNull(oRecordSet!Category), "", oRecordSet!Category), MResourse.ENCRYPT_KEY)
      oRecordSet!userName = goBlowfish.EncryptString(IIf(IsNull(oRecordSet!userName), "", oRecordSet!userName), MResourse.ENCRYPT_KEY)
      oRecordSet![password] = goBlowfish.EncryptString(IIf(IsNull(oRecordSet![password]), "", oRecordSet![password]), MResourse.ENCRYPT_KEY)
      oRecordSet.Update
      oRecordSet.MoveNext
      
      'Move progress bar.
      Call goControl.ChangeProgressBarValue(pbrEncrypt, 100, recordCount, i)
   Wend

   pbrEncrypt.Value = 0
   oRecordSet.Close
End Sub

Public Sub DecryptDB()
   Dim oRecordSet As New ADODB.Recordset
   Dim sSql As String
   Dim strFieldValue
   Dim recordCount, i As Double
   i = 1
   
   
   sSql = "SELECT * FROM Links"
   Call oRecordSet.Open(sSql, goAdo.GetConnectionAccess(gDBName), adOpenStatic, adLockOptimistic)
  
   recordCount = oRecordSet.recordCount

   While oRecordSet.EOF <> True
   
      oRecordSet!Title = goBlowfish.DecryptString(IIf(IsNull(oRecordSet!Title), "", oRecordSet!Title), MResourse.ENCRYPT_KEY)
      oRecordSet!Address = goBlowfish.DecryptString(IIf(IsNull(oRecordSet!Address), "", oRecordSet!Address), MResourse.ENCRYPT_KEY)
      oRecordSet!Category = goBlowfish.DecryptString(IIf(IsNull(oRecordSet!Category), "", oRecordSet!Category), MResourse.ENCRYPT_KEY)
      oRecordSet!userName = goBlowfish.DecryptString(IIf(IsNull(oRecordSet!userName), "", oRecordSet!userName), MResourse.ENCRYPT_KEY)
      oRecordSet![password] = goBlowfish.DecryptString(IIf(IsNull(oRecordSet![password]), "", oRecordSet![password]), MResourse.ENCRYPT_KEY)
      oRecordSet.Update
      oRecordSet.MoveNext
      
      'Move progress bar.
      Call goControl.ChangeProgressBarValue(pbrEncrypt, 100, recordCount, i)
   Wend

   pbrEncrypt.Value = 0
   oRecordSet.Close
End Sub

Public Sub FillDataGrid()
   Dim sSql As String
   
   'Fill datagrid.
   sSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount  FROM Links ORDER BY Title"
   Call goControl.FillMSHFlexGrid(fgridAddresses, gDBName, goRecordset, sSql)
End Sub

Public Sub BrowseDataGrid()
   Call goControl.FillMSHFlexGrid(fgridAddresses, gDBName, goRecordset, goBrowseSql)
End Sub

Private Function ControlSameLink(pAddress As String) As Boolean
   Dim sSql As String
   Dim oRecordSet As New Recordset
   
   sSql = "SELECT Address FROM Links WHERE Address = '" & pAddress & "'"
   Set oRecordSet = goAdo.GetRecordSetByFileName(gDBName, sSql)
      
   If (oRecordSet.recordCount = 0) Then
      ControlSameLink = False
   Else
      ControlSameLink = True
   End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''
'SSTab 2 Codes
''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''
'''MENU Members'''
''''''''''''''''''
Private Sub mnBrowse_Click()
    'Same code with fgridAddresses_DblClick()

   Dim strAddress, strCmdArg, sSql As String
      
   'If a header column is dblclicked then return
   If (ggridMouseRow = 0) Then
       Exit Sub
   End If
   
       
   'Get the internet link
   strAddress = fgridAddresses.TextMatrix(ggridMouseRow, 1)
   Me.Caption = strAddress


   'Goto web site
   BrowseWebSite (strAddress)


   'Update the global recordset, increase VisitCount & update LastDate
   sSql = "UPDATE Links SET VisitCount = VisitCount + 1, LastDate = '" & DateTime.Now() & "' WHERE Address ='" & strAddress & "';"
   Call goAdo.ExecuteCommand(gDBName, sSql)
   
   
   'Fill or browse datagrid
   If goBrowseSql = "" Then
      FillDataGrid
   Else
      BrowseDataGrid
   End If
    
    
   'Add item to ComboBox from which the search is done.
   AddItemToComboBox
End Sub

Private Sub mnEdit_Click()
   Dim strLink As String
   Dim strTitle As String
   Dim strUserName As String
   Dim strPassword As String
      
   strLink = fgridAddresses.TextMatrix(fgridAddresses.Row, 1)
   strTitle = fgridAddresses.TextMatrix(fgridAddresses.Row, 0)
   strCategory = fgridAddresses.TextMatrix(fgridAddresses.Row, 2)
   strUserName = fgridAddresses.TextMatrix(fgridAddresses.Row, 3)
   strPassword = fgridAddresses.TextMatrix(fgridAddresses.Row, 4)
   
   'Fill edit form data
   frmEditLink.txtTitle.Text = strTitle
   frmEditLink.txtLink.Text = strLink
   frmEditLink.txtUserName.Text = strUserName
   frmEditLink.txtPassword.Text = strPassword
      
   'Fill categories & select the category of the data
   Call frmEditLink.cmbCategory.FillComboBoxFromComboBox(cmbCategory)
   frmEditLink.cmbCategory.SeletedValue = strCategory
   frmEditLink.Show (vbModal)
End Sub

Private Sub mnDelete_Click()
   Dim strAddress, strTitle As String
   strTitle = fgridAddresses.TextMatrix(fgridAddresses.Row, 0)
   strAddress = fgridAddresses.TextMatrix(fgridAddresses.Row, 1)
   
   
   rMsg = MsgBox("Do you want to delete " & strTitle & " link ?", vbYesNo + vbQuestion, MApplication.ProductNameVersion + " - Delete confirmation")
   If rMsg = vbNo Then
      Exit Sub
   End If
   
   'Delete selected record
   sSql = "DELETE FROM Links WHERE Address ='" & strAddress & "';"
   Call goAdo.ExecuteCommand(gDBName, sSql)

   
   'Fill or browse datagrid
   If goBrowseSql = "" Then
      FillDataGrid
   Else
      BrowseDataGrid
   End If

End Sub

Private Sub cmbBrowseTitle_Change()
   'Set cmbIndex for AddItemToComboBox
   cmbIndex = BrowseTitle
   
   gtitleKey = Trim(cmbBrowseTitle.Text)
   Call InitializeSearchCriteria
     
   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category LIKE '%" & gcategoryKey & "%' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   BrowseDataGrid
End Sub

Private Sub cmbBrowseTitle_Click()
   gtitleKey = cmbBrowseTitle.List(cmbBrowseTitle.ListIndex)
   Call InitializeSearchCriteria

   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category LIKE '%" & gcategoryKey & "%' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   BrowseDataGrid
End Sub

Private Sub cmbBrowseAddress_Change()
   'Set cmbIndex for AddItemToComboBox
   cmbIndex = BrowseAddress
   
   gaddressKey = Trim(cmbBrowseAddress.Text)
   Call InitializeSearchCriteria
   
   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category LIKE '%" & gcategoryKey & "%' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   BrowseDataGrid
End Sub

Private Sub cmbBrowseAddress_Click()
   gaddressKey = cmbBrowseAddress.List(cmbBrowseAddress.ListIndex)
   Call InitializeSearchCriteria

   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category LIKE '%" & gcategoryKey & "%' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   BrowseDataGrid
End Sub

Private Sub btnDeleteLink_Click()
   cmbBrowseAddress.Clear
End Sub

Private Sub btnDeleteTitle_Click()
   cmbBrowseTitle.Clear
End Sub

Private Sub cmbBrowseCat_Change()
   gcategoryKey = Trim(cmbBrowseCat.Text)
   Call InitializeSearchCriteria
   
   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category LIKE '%" & gcategoryKey & "%' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   
    'Fill or browse datagrid
   If goBrowseSql = "" Then
      FillDataGrid
   Else
      BrowseDataGrid
   End If

End Sub

Private Sub cmbBrowseCat_Click()
   gcategoryKey = cmbBrowseCat.List(cmbBrowseCat.ListIndex)
   Call InitializeSearchCriteria

   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category = '" & gcategoryKey & "' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   
    'Fill or browse datagrid
   If goBrowseSql = "" Then
      FillDataGrid
   Else
      BrowseDataGrid
   End If

End Sub

Private Sub txtBrowsePassword_Change()
   gpasswordKey = Trim(txtBrowsePassword.Text)
   Call InitializeSearchCriteria

   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category LIKE '%" & gcategoryKey & "%' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   BrowseDataGrid
End Sub

Private Sub txtBrowseUserName_Change()
   guserNameKey = Trim(txtBrowseUserName.Text)
   Call InitializeSearchCriteria
      
   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE Title LIKE '%" & gtitleKey & "%' AND Address LIKE '%" & gaddressKey & "%' AND Category LIKE '%" & gcategoryKey & "%' AND UserName LIKE '%" & guserNameKey & "%' AND [Password] LIKE '%" & gpasswordKey & "%' ORDER BY Title"
   BrowseDataGrid
End Sub

Private Sub dtLastVisit_Change()
   Dim Key As String
   Dim stoday As String
   Key = Month(dtLastVisit.Value) & "/" & Day(dtLastVisit.Value) & "/" & Year(dtLastVisit.Value)
   stoday = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
         
   goBrowseSql = "SELECT Title, Address, Category, UserName, [Password], LastDate, FirstDate, VisitCount FROM Links WHERE LastDate BETWEEN #" & stoday & "# AND #" & Key & "# ORDER BY Title"
   BrowseDataGrid
End Sub

Private Sub InitializeSearchCriteria()
   
   'If Title is an empty string then find all Titles
   If (gtitleKey = "") Then
      gtitleKey = "%"
   End If
   
   'If Address is an empty string then find all Addresses.
   If (gaddressKey = "") Then
      gaddressKey = "%"
   End If
   
   'If any Category was not selected then find in all Categories
   If (gcategoryKey = "") Then
      gcategoryKey = "%"
   End If
   
   'If key is an empty string then find all records.
   If (guserNameKey = "") Then
      guserNameKey = "%"
   End If
   
   'If key is an empty string then find all records.
   If (gpasswordKey = "") Then
      gpasswordKey = "%"
   End If
   
   
   
End Sub

Private Sub btnAll_Click()
   cmbBrowseTitle.Text = ""
   cmbBrowseAddress.Text = ""
   txtBrowseUserName.Text = ""
   txtBrowsePassword.Text = ""
   cmbBrowseCat.Text = ""
   
   'Select all categories
   'cmbBrowseCat.ListIndex = -1
   
   'Call InitializeSearchCriteria
   
   'Fill datagrid
   FillDataGrid

End Sub

Private Sub btnSaveLinks_Click()
   Dim strHeader As String
   Dim strFooter As String
   Dim strTableRowTemplateFirst As String
   Dim strTableRowTemplateSecond As String
   Dim strTableRow As String
   Dim strIndex As String
   
   Dim strTitle As String
   Dim strAddress As String
   Dim strCategory As String
   Dim strUserName As String
   Dim strPassword As String
   Dim strFirstVisitDate As String
   Dim strLastVisitDate As String
   Dim strTotalVisit As String
   Dim strSaveFilePath As String
   
   
   'Open save file dialog
   strSaveFilePath = goDialog.OpenSaveDialog(dlgSaveFile, "Save File Into...", "HTML Files (*.html, *.htm)|*.htm;*.html", "*.htm")
      
   'Construct Index
   strHeader = MResourse.htmlHeader
   strTableRowTemplateFirst = MResourse.htmlTableRowTemplateFirst
   strTableRowTemplateSecond = MResourse.htmlTableRowTemplateSecond
   strFooter = MResourse.htmlFooter
  
   'Append header
   strHeader = Replace(strHeader, "$DocumentTitle$", strSaveFilePath)
   strIndex = strHeader
   
   
   For i = 1 To fgridAddresses.Rows - 1
      
      'Get column values of a row
      If (chkTitle.Value) Then
         strTitle = fgridAddresses.TextMatrix(i, 0)
      Else
         strTitle = ""
      End If
      
      If (chkAddress.Value) Then
         strAddress = fgridAddresses.TextMatrix(i, 1)
      Else
         strAddress = ""
      End If
      
      If (chkCategory.Value) Then
         strCategory = fgridAddresses.TextMatrix(i, 2)
      Else
         strCategory = ""
      End If
      
      If (chkUserName.Value) Then
         strUserName = fgridAddresses.TextMatrix(i, 3)
      Else
         strUserName = ""
      End If
      
      If (chkPassword.Value) Then
         strPassword = fgridAddresses.TextMatrix(i, 4)
      Else
         strPassword = ""
      End If
      
      If (chkFirstVisitDate.Value) Then
         strFirstVisitDate = fgridAddresses.TextMatrix(i, 5)
      Else
         strFirstVisitDate = ""
      End If
      
      If (chkLastVisitDate.Value) Then
         strLastVisitDate = fgridAddresses.TextMatrix(i, 6)
      Else
         strLastVisitDate = ""
      End If
      
      If (chkTotalVisit.Value) Then
         strTotalVisit = fgridAddresses.TextMatrix(i, 7)
      Else
         strTotalVisit = ""
      End If
      
      
      'Set alternating rows
      If (i Mod 2 = 0) Then
         strTableRow = strTableRowTemplateSecond
      Else
         strTableRow = strTableRowTemplateFirst
      End If
      
      
      'Replace templates
      strTableRow = Replace(strTableRow, "$Number$", i)
      strTableRow = Replace(strTableRow, "$Title$", strTitle)
      strTableRow = Replace(strTableRow, "$Address$", strAddress)
      strTableRow = Replace(strTableRow, "$Category$", strCategory)
      strTableRow = Replace(strTableRow, "$UserName$", strUserName)
      strTableRow = Replace(strTableRow, "$Password$", strPassword)
      strTableRow = Replace(strTableRow, "$FirstVisitDate$", strFirstVisitDate)
      strTableRow = Replace(strTableRow, "$LastVisitDate$", strLastVisitDate)
      strTableRow = Replace(strTableRow, "$TotalVisit$", strTotalVisit)
      
      
      'Append row to main body
      strIndex = strIndex & strTableRow & vbNewLine
   Next
   
   
   'Append footer
   strIndex = strIndex & strFooter
      
   'Save file
   Call goIO.WriteFile(strIndex, strSaveFilePath)
   
End Sub

Private Sub fgridAddresses_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Store the mouse button
   If Button = MouseButtonConstants.vbLeftButton Then
      gmouseButton = vbLeftButton
   ElseIf Button = MouseButtonConstants.vbRightButton Then
      gmouseButton = vbRightButton
   End If
End Sub

Private Sub fgridAddresses_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Get the MouseRow of the grid
   ggridMouseRow = fgridAddresses.MouseRow
End Sub

Private Sub fgridAddresses_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim sSql, strAddress As String
   
   
   If KeyCode = KeyCodeConstants.vbKeyDelete Then
   
      'Get the link of the active selected row
      strAddress = fgridAddresses.TextMatrix(fgridAddresses.Row, 1)
      
      'Get the title of the active selected row
      strTitle = fgridAddresses.TextMatrix(fgridAddresses.Row, 0)
      
      rMsg = MsgBox("Do you want to delete " + strTitle, vbYesNo + vbQuestion, MApplication.ProductNameVersion + " - Delete confirmation")
      If rMsg = vbNo Then
         Exit Sub
      End If
      
      'Delete selected record
      sSql = "DELETE FROM Links WHERE Address ='" & strAddress & "';"
      Call goAdo.ExecuteCommand(gDBName, sSql)
      
      
      'Fill or browse datagrid
      If goBrowseSql = "" Then
         FillDataGrid
      Else
         BrowseDataGrid
      End If
      
      
   
   ElseIf KeyCode = KeyCodeConstants.vbKeyReturn Then
            
            
      'Get the internet link
      strAddress = fgridAddresses.TextMatrix(fgridAddresses.Row, 1)
      Me.Caption = strAddress


      'Goto web site
      BrowseWebSite (strAddress)
      
      
      'Update the global recordset, increase VisitCount & update LastDate
      sSql = "UPDATE Links SET VisitCount = VisitCount + 1, LastDate = '" & DateTime.Now() & "' WHERE Address ='" & strAddress & "';"
      Call goAdo.ExecuteCommand(gDBName, sSql)
      
      
      'Fill or browse datagrid
      If goBrowseSql = "" Then
         FillDataGrid
      Else
         BrowseDataGrid
      End If
    
    
      'Add item to ComboBox from which the search is done.
      AddItemToComboBox
   
   End If
   
   
End Sub

Private Sub fgridAddresses_Click()

   Dim sfieldName As String
   fgridAddresses.SetFocus
   
   If gmouseButton = vbLeftButton Or gmouseButton = vbRightButton Then
     
     'If the first row is clicked then sort the grid.
     If (fgridAddresses.MouseRow = 0) Then
         Select Case fgridAddresses.MouseCol
            'Type of the first 2 fields are memo.
            'Memo type cannot be sorted in fieldname consept.
            Case 0:
               Call goControl.SortMSHFlexGrid(fgridAddresses)
            Case 1:
               Call goControl.SortMSHFlexGrid(fgridAddresses)
            Case 2:
               sfieldName = "Category"
               Call goControl.SortMSHFlexGridByField(fgridAddresses, sfieldName)
            Case 3:
               sfieldName = "UserName"
               Call goControl.SortMSHFlexGridByField(fgridAddresses, sfieldName)
            Case 4:
               sfieldName = "Password"
               Call goControl.SortMSHFlexGridByField(fgridAddresses, sfieldName)
            Case 5:
               sfieldName = "LastDate"
               Call goControl.SortMSHFlexGridByField(fgridAddresses, sfieldName)
            Case 6:
               sfieldName = "FirstDate"
               Call goControl.SortMSHFlexGridByField(fgridAddresses, sfieldName)
            Case 7:
               sfieldName = "VisitCount"
               Call goControl.SortMSHFlexGridByField(fgridAddresses, sfieldName)
         End Select
     End If
     
  End If
  
  
  'When the right mouse button is clicked open a popup menu
  If gmouseButton = vbRightButton Then
      
      'If the first row is not clicked then open popup
      If (fgridAddresses.MouseRow <> 0) Then
         Call goDialog.OpenPopupDialog(Me, mnGridPopup)
      End If
   End If
  
End Sub

Private Sub fgridAddresses_DblClick()
   Dim strAddress, strCmdArg, sSql As String
      
   'If a header column is dblclicked then return
   If (ggridMouseRow = 0) Then
       Exit Sub
   End If
   
       
   'Get the internet link
   strAddress = fgridAddresses.TextMatrix(ggridMouseRow, 1)
   Me.Caption = strAddress


   'Goto web site
   BrowseWebSite (strAddress)


   'Update the global recordset, increase VisitCount & update LastDate
   sSql = "UPDATE Links SET VisitCount = VisitCount + 1, LastDate = '" & DateTime.Now() & "' WHERE Address ='" & strAddress & "';"
   Call goAdo.ExecuteCommand(gDBName, sSql)
   
   
   'Fill or browse datagrid
   If goBrowseSql = "" Then
      FillDataGrid
   Else
      BrowseDataGrid
   End If
    
    
   'Add item to ComboBox from which the search is done.
   AddItemToComboBox

End Sub

Private Sub AddItemToComboBox()
   'Adds last searched item to a combobox
   
   If (cmbIndex = ComboBoxName.BrowseTitle) Then
      Call goControl.AddItemToComboBoxUnique(cmbBrowseTitle, LCase(cmbBrowseTitle.Text))
   ElseIf (cmbIndex = ComboBoxName.BrowseAddress) Then
      Call goControl.AddItemToComboBoxUnique(cmbBrowseAddress, LCase(cmbBrowseAddress.Text))
   End If
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMonitor_Timer()
    Dim url As String
    Dim Title As String

    'Choose browser type
    If rbtnFireFox.Value Then
      Call GetItemFromFirefox(url, Title)
    ElseIf rbtnIExplorer.Value Then
      Call GetItemFromIE(url, Title)
    End If
    
    txtInfo.Text = Title
    txtAddress.Text = url
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub FillcmbBrowseCat()
   Dim sSql As String
   
   'First clear combobox
   cmbBrowseCat.Clear
   
   'Fill ComboBox
   sSql = "SELECT DISTINCT Category FROM Links"
   Call goControl.FillComboBox(cmbBrowseCat, gDBName, goRecordset, sSql)
End Sub

Private Sub tmrRowCount_Timer()
   'Count rows of datagrid - 1
   'Total # records
   sbarRowCount.Panels(3).Text = MResourse.rTotal & " " & fgridAddresses.Rows - 1 & " " & MResourse.rRecord
End Sub

Private Sub BrowseWebSite(ByVal plink As String)

   Dim strCmdArg As String
   
   If (rbtnFireFoxNav.Value) Then
   
      'Open the selected link with Mozilla Firefox
      strCmdArg = "C:\Program Files\Mozilla Firefox\firefox.exe " & plink
      Call Shell(strCmdArg, vbNormalFocus)
   ElseIf (rbtnIExplorerNav.Value) Then
   
      'Open the selected link with Internet Explorer
      strCmdArg = "c:\program files\internet explorer\iexplore.exe " & plink
      Call Shell(strCmdArg, vbNormalFocus)
   End If
End Sub

Private Sub GetItemFromIE(purl As String, ptitle As String)
    'Get Title & URL of the active explorer
    Call goCLink.LinkTextBox(txtMonitor, "iexplore|WWW_GetWindowInfo", vbLinkManual, &HFFFFFFFF)
        
   Dim sTmp As String, sLink() As String

    sTmp = Mid(txtMonitor.Text, 2)
    sLink = Split(sTmp, """,""")
    
    If UBound(sLink) > 0 Then
        If left(sLink(0), 1) = """" Then
            sTmp = Mid(sLink(0), 2, Len(sLink(0)) - 1)
        Else
            sTmp = sLink(0)
        End If

            purl = sTmp

            If sLink(1) <> "" Then
                If Right(sLink(1), 1) = """" Then
                    sTmp = left(sLink(1), Len(sLink(1)) - 1)
                Else
                    sTmp = sLink(1)
                End If
                    
                ptitle = sTmp
            End If
        
    End If
End Sub

Private Sub GetItemFromFirefox(purl As String, ptitle As String)
    'Get Title & URL of the active explorer
    Call goCLink.LinkTextBox(txtMonitor, "firefox|WWW_GetWindowInfo", vbLinkManual, &HFFFFFFFF)
        
    Dim sTmp As String, sLink() As String

    sTmp = Mid(txtMonitor.Text, 2)
    sLink = Split(sTmp, """,""")
    
    If UBound(sLink) > 0 Then
        If left(sLink(0), 1) = """" Then
            sTmp = Mid(sLink(0), 2, Len(sLink(0)) - 1)
        Else
            sTmp = sLink(0)
        End If

        purl = sTmp
 
            If sLink(1) <> "" Then
                If Right(sLink(1), 1) = """" Then
                    sTmp = left(sLink(1), Len(sLink(1)) - 1)
                Else
                    sTmp = sLink(1)
                End If
                    
                ptitle = sTmp
            End If
        
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''
'SSTab 3 Codes
''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub fgridFavAddresses_DblClick()
   Dim strAddress, strCmdArg, sSql As String
      
   'If a header column is dblclicked then return
   If (ggridMouseRow = 0) Then
       Exit Sub
   End If
   
       
   'Get the internet link
   strAddress = fgridFavAddresses.TextMatrix(ggridMouseRow, 1)
   Me.Caption = strAddress


   'Goto web site
   BrowseWebSite (strAddress)


   'Update the global recordset, increase VisitCount & update LastDate
   sSql = "UPDATE Links SET VisitCount = VisitCount + 1, LastDate = '" & DateTime.Now() & "' WHERE Address ='" & strAddress & "';"
   Call goAdo.ExecuteCommand(gDBName, sSql)
   
   
   'Fill or browse datagrid
   If goBrowseSql = "" Then
      FillDataGrid
   Else
      BrowseDataGrid
   End If

End Sub

Private Sub fgridFavAddresses_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sSql, strAddress As String
   
   
   If KeyCode = KeyCodeConstants.vbKeyDelete Then
   
      'Get the link of the active selected row
      strAddress = fgridFavAddresses.TextMatrix(fgridFavAddresses.Row, 1)
      
      'Get the title of the active selected row
      strTitle = fgridFavAddresses.TextMatrix(fgridFavAddresses.Row, 0)
      
      rMsg = MsgBox("Do you want to delete " + strTitle, vbYesNo + vbQuestion, MApplication.ProductNameVersion + " - Delete confirmation")
      If rMsg = vbNo Then
         Exit Sub
      End If
      
      'Delete selected record
      sSql = "DELETE FROM Links WHERE Address ='" & strAddress & "';"
      Call goAdo.ExecuteCommand(gDBName, sSql)
      
      
      'Fill or browse datagrid
      If goBrowseSql = "" Then
         FillDataGrid
      Else
         BrowseDataGrid
      End If
      
      
   
   ElseIf KeyCode = KeyCodeConstants.vbKeyReturn Then
            
            
      'Get the internet link
      strAddress = fgridFavAddresses.TextMatrix(fgridFavAddresses.Row, 1)
      Me.Caption = strAddress


      'Goto web site
      BrowseWebSite (strAddress)
      
      
      'Update the global recordset, increase VisitCount & update LastDate
      sSql = "UPDATE Links SET VisitCount = VisitCount + 1, LastDate = '" & DateTime.Now() & "' WHERE Address ='" & strAddress & "';"
      Call goAdo.ExecuteCommand(gDBName, sSql)
      
      
      'Fill or browse datagrid
      If goBrowseSql = "" Then
         FillDataGrid
      Else
         BrowseDataGrid
      End If
   
   End If
   
   
End Sub

Private Sub btnBrowseLinks_Click()
   Dim strAddress As String
   Dim sSql As String
           
         
   'Browse all links in the grid
   For i = 1 To fgridFavAddresses.Rows - 1
      strAddress = fgridFavAddresses.TextMatrix(i, 1)
      BrowseWebSite (strAddress)
   Next
   
   
   'Update all visited records
   For i = 1 To fgridFavAddresses.Rows - 1
      strAddress = fgridFavAddresses.TextMatrix(i, 1)
      
      'Update the global recordset, increase VisitCount & update LastDate
      sSql = "UPDATE Links SET VisitCount = VisitCount + 1, LastDate = '" & DateTime.Now() & "' WHERE Address ='" & strAddress & "';"
      Call goAdo.ExecuteCommand(gDBName, sSql)
   Next
End Sub

Private Sub btnViewLinks_Click()
   LoadFavoriteLinks (cmbFavoriteLinks.Text)
End Sub
