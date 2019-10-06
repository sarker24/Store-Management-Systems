VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmStock 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Receiving  Informations [RCH Mother Store]"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReceiving.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAddItem 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Input Item"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13800
      Picture         =   "frmReceiving.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton cmdUndoPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Undo Post"
      Height          =   855
      Left            =   11040
      Picture         =   "frmReceiving.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9720
      Width           =   1215
   End
   Begin VB.CommandButton CmdPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Post"
      Height          =   855
      Left            =   10080
      Picture         =   "frmReceiving.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   855
      Left            =   9120
      Picture         =   "frmReceiving.frx":1FE8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      Height          =   855
      Left            =   8160
      Picture         =   "frmReceiving.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Print"
      Height          =   855
      Left            =   7080
      Picture         =   "frmReceiving.frx":317C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9720
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Q&uit"
      Height          =   855
      Left            =   6120
      Picture         =   "frmReceiving.frx":3A46
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   855
      Left            =   5160
      Picture         =   "frmReceiving.frx":4310
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      Height          =   855
      Left            =   4200
      Picture         =   "frmReceiving.frx":4BDA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   855
      Left            =   3240
      Picture         =   "frmReceiving.frx":54A4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   855
      Left            =   2280
      Picture         =   "frmReceiving.frx":5D6E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9720
      Width           =   990
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Find Last"
      Top             =   10080
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Find Next"
      Top             =   10080
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Find Previous"
      Top             =   9720
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Find First"
      Top             =   9720
      Width           =   735
   End
   Begin VB.TextBox txtUName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   9720
      Width           =   1335
   End
   Begin VB.CommandButton cmdLDelete 
      BackColor       =   &H00C0B4A9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   15600
      Picture         =   "frmReceiving.frx":6638
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3120
      Width           =   420
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Receiving Details Information"
      ForeColor       =   &H00008000&
      Height          =   7455
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   14655
      Begin VB.CommandButton cmdRDelete 
         BackColor       =   &H00C0B4A9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14160
         Picture         =   "frmReceiving.frx":6BC2
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   420
      End
      Begin VSFlex7DAOCtl.VSFlexGrid fgStock 
         Height          =   6855
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   14295
         _cx             =   25215
         _cy             =   12091
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   0
         BackColorFixed  =   0
         ForeColorFixed  =   49152
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483634
         BackColorBkg    =   12629161
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmReceiving.frx":714C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Receiving Master Information"
      ForeColor       =   &H00008000&
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   14655
      Begin VB.ComboBox cmbSName 
         Height          =   360
         Left            =   5040
         TabIndex        =   36
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtTotalBill 
         Height          =   375
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtPaid 
         Height          =   375
         Left            =   6960
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtDue 
         Height          =   375
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtpost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtSerialNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1560
         TabIndex        =   18
         Text            =   " "
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtSBill 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9240
         TabIndex        =   1
         Text            =   " "
         Top             =   360
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc dcCatagory 
         Height          =   720
         Left            =   9480
         Top             =   6720
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   1270
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dcItemGroup"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker RDate 
         Height          =   405
         Left            =   1560
         TabIndex        =   20
         Top             =   360
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65273859
         CurrentDate     =   37840
      End
      Begin VB.Label lblStore 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Store Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11400
         TabIndex        =   39
         Top             =   840
         Width           =   1095
      End
      Begin MSForms.ComboBox cmbStoreName 
         Height          =   375
         Left            =   12480
         TabIndex        =   38
         Top             =   840
         Width           =   2055
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3625;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbDepartment 
         Height          =   375
         Left            =   12480
         TabIndex        =   37
         Top             =   360
         Width           =   2055
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3625;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label lblDue 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Due"
         Height          =   375
         Left            =   8160
         TabIndex        =   34
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblPaid 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Paid"
         Height          =   375
         Left            =   6360
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblTotalBill 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Bill"
         Height          =   375
         Left            =   3840
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblSupplierBill 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Supplier Bill"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblChallanNO 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Challan No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblBudgetHead 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Department Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10920
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblSupplier 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Supplier Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Receiving Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   600
      Left            =   -480
      Top             =   12720
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1058
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dcItemGroup"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   480
      Top             =   10560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Driver={SQL Server};Server=MAS;Database=KG;Trusted_Connection=yes"
      OLEDBString     =   "Driver={SQL Server};Server=MAS;Database=KG;Trusted_Connection=yes"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblCashMaster"
      Caption         =   "Record Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INVENTORY ITEM PURCHASE INFORMATION"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 Private rsItemMaster                    As ADODB.Recordset
 Private rsItemDetail                    As ADODB.Recordset
 Private rsDMaster                      As New ADODB.Recordset
 Private rsDDetail                      As ADODB.Recordset
 Private rs                              As ADODB.Recordset
 Private rsTemp2                         As ADODB.Recordset

 Private bRecordExists                   As Boolean
 Dim str As String
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'----Add For Reporting Perpose----------------------------------------------
Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition


Private objReportSub                        As CRPEAuto.Report 'sub
Private objReportDatabaseSub                As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition

                          
Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private rsDailyRpt                          As ADODB.Recordset
Private Tracer                              As Integer
Private strGroupName                        As String
Dim temp As Double
Dim temp1 As Double
Dim temp2 As Double
'Private temp As Double
'Private temp1 As Double
'--------------------------------------------------------------------------------

Private Sub cmbSName_GotFocus()
cmbSName.BackColor = &HFFFFC0
End Sub

Private Sub cmbSName_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(cmbSName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbSName_LostFocus()
    cmbSName.BackColor = vbWhite
End Sub

Private Sub SName()
    Dim rsTemp2 As New ADODB.Recordset
          
     rsTemp2.Open ("SELECT DISTINCT sName FROM SSuplierName ORDER BY sName ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbSName.AddItem rsTemp2("sName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub


Private Sub cmdFirst_Click()
Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset

Adodc2.Recordset.MoveFirst
If Adodc2.Recordset.EOF = True Then
          MsgBox "end of file"
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    txtSerialNo = rsItemMaster!SerialNo
    RDate = rsItemMaster!RDate
    cmbSName = rsItemMaster!SName
    txtSBill = rsItemMaster!SBill
    cmbDepartment = rsItemMaster!DName
    cmbStoreName = rsItemMaster!StoreName
    txtTotalBill = rsItemMaster!TotalBill
    txtPaid = rsItemMaster!Paid
    txtDue = rsItemMaster!Due
    txtpost = rsItemMaster!Posted
    txtUName = rsItemMaster!UName


    fgStock.Rows = 1
    strPaymentDetail = "SELECT  SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit " & _
                      "FROM SStockDetails " & _
                      "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly

 If rsItemDetail.RecordCount <> 0 Then

        fgStock.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgStock.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgStock.TextMatrix(i, 2) = rsItemDetail("RDate")
            fgStock.TextMatrix(i, 3) = rsItemDetail("SName")
            fgStock.TextMatrix(i, 4) = rsItemDetail("StoreName")
            fgStock.TextMatrix(i, 5) = rsItemDetail("Catagory")
            fgStock.TextMatrix(i, 6) = rsItemDetail("ItemName")
            fgStock.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgStock.TextMatrix(i, 8) = rsItemDetail("Rate")
            fgStock.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgStock.TextMatrix(i, 10) = rsItemDetail("Rol")
            fgStock.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("ExpDate")), "", Format(rsItemDetail("ExpDate"), "dd-mmm-yyyy"))
            fgStock.TextMatrix(i, 12) = rsItemDetail("Posted")
            fgStock.TextMatrix(i, 13) = rsItemDetail("Warranty")
            fgStock.TextMatrix(i, 14) = rsItemDetail("Remarks")
            fgStock.TextMatrix(i, 15) = rsItemDetail("CPost")
            fgStock.TextMatrix(i, 16) = rsItemDetail("Unit")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close

        
        
End If
End Sub

Private Sub cmdLast_Click()
Dim i As Integer
    Dim strLedgerDetail As String
    Set rsItemDetail = New ADODB.Recordset

Adodc2.Recordset.MoveLast
If Adodc2.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    txtSerialNo = rsItemMaster!SerialNo
    RDate = rsItemMaster!RDate
    cmbSName = rsItemMaster!SName
    txtSBill = rsItemMaster!SBill
    cmbDepartment = rsItemMaster!DName
    cmbStoreName = rsItemMaster!StoreName
    txtTotalBill = rsItemMaster!TotalBill
    txtPaid = rsItemMaster!Paid
    txtDue = rsItemMaster!Due
    txtpost = rsItemMaster!Posted
    txtUName = rsItemMaster!UName
        
        
    fgStock.Rows = 1
    strLedgerDetail = "SELECT  SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit " & _
                      "FROM SStockDetails " & _
                      "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsItemDetail.RecordCount <> 0 Then

        fgStock.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgStock.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgStock.TextMatrix(i, 2) = rsItemDetail("RDate")
            fgStock.TextMatrix(i, 3) = rsItemDetail("SName")
            fgStock.TextMatrix(i, 4) = rsItemDetail("StoreName")
            fgStock.TextMatrix(i, 5) = rsItemDetail("Catagory")
            fgStock.TextMatrix(i, 6) = rsItemDetail("ItemName")
            fgStock.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgStock.TextMatrix(i, 8) = rsItemDetail("Rate")
            fgStock.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgStock.TextMatrix(i, 10) = rsItemDetail("Rol")
            fgStock.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("ExpDate")), "", Format(rsItemDetail("ExpDate"), "dd-mmm-yyyy"))
            fgStock.TextMatrix(i, 12) = rsItemDetail("Posted")
            fgStock.TextMatrix(i, 13) = rsItemDetail("Warranty")
            fgStock.TextMatrix(i, 14) = rsItemDetail("Remarks")
            fgStock.TextMatrix(i, 15) = rsItemDetail("CPost")
            fgStock.TextMatrix(i, 16) = rsItemDetail("Unit")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close

        
        
End If
End Sub

Private Sub cmdNext_Click()
Dim i As Integer
    Dim strLedgerDetail As String
    Set rsItemDetail = New ADODB.Recordset
    
Adodc2.Recordset.MoveNext
If Adodc2.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    txtSerialNo = rsItemMaster!SerialNo
    RDate = rsItemMaster!RDate
    cmbSName = rsItemMaster!SName
    txtSBill = rsItemMaster!SBill
    cmbDepartment = rsItemMaster!DName
    cmbStoreName = rsItemMaster!StoreName
    txtTotalBill = rsItemMaster!TotalBill
    txtPaid = rsItemMaster!Paid
    txtDue = rsItemMaster!Due
    txtpost = rsItemMaster!Posted
    txtUName = rsItemMaster!UName
        
        
    fgStock.Rows = 1
    
    strLedgerDetail = "SELECT  SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit " & _
                      "FROM SStockDetails " & _
                      "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "
    
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsItemDetail.RecordCount <> 0 Then

        fgStock.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgStock.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgStock.TextMatrix(i, 2) = rsItemDetail("RDate")
            fgStock.TextMatrix(i, 3) = rsItemDetail("SName")
            fgStock.TextMatrix(i, 4) = rsItemDetail("StoreName")
            fgStock.TextMatrix(i, 5) = rsItemDetail("Catagory")
            fgStock.TextMatrix(i, 6) = rsItemDetail("ItemName")
            fgStock.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgStock.TextMatrix(i, 8) = rsItemDetail("Rate")
            fgStock.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgStock.TextMatrix(i, 10) = rsItemDetail("Rol")
            fgStock.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("ExpDate")), "", Format(rsItemDetail("ExpDate"), "dd-mmm-yyyy"))
            fgStock.TextMatrix(i, 12) = rsItemDetail("Posted")
            fgStock.TextMatrix(i, 13) = rsItemDetail("Warranty")
            fgStock.TextMatrix(i, 14) = rsItemDetail("Remarks")
            fgStock.TextMatrix(i, 15) = rsItemDetail("CPost")
            fgStock.TextMatrix(i, 16) = rsItemDetail("Unit")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
    End If
End Sub

Private Sub cmdPreview_Click()
    Call printReport
End Sub


Private Sub cmdAddItem_Click()
frmItemInputReceiving.Show vbModal
Call Calculation
End Sub


Private Sub postedCheck()
      Dim f As Integer
      Dim i As Integer
      f = fgStock.Rows - 1
'      If chkAutoposting.Value = 1 Then
      
For i = 1 To f
    fgStock.Cell(flexcpChecked, i, 12) = flexChecked
Next i


End Sub
Private Sub cmdCancel_Click()
    
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "select SerialNo,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & frmLogin.txtUID.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
            
 If rs!Privilegegroup = 0 Then
    cmdCancel.Enabled = False
    CmdNew.Enabled = True
    CmdEdit.Caption = "&Edit"
    CmdNew.Caption = "&New"
    cmdClose.Enabled = True
    CmdEdit.Enabled = True
    CmdOpen.Enabled = True
    cmdPost.Caption = "&Post"
'    cmdDelete.Enabled = True
    cmdPreview.Enabled = True
    cmdPost.Enabled = True
    Call alldisable
'    If Not rsItemMaster.EOF Then FindRecord
Else
cmdCancel.Enabled = False
    CmdNew.Enabled = True
    CmdEdit.Caption = "&Edit"
    CmdNew.Caption = "&New"
    cmdClose.Enabled = True
    CmdEdit.Enabled = True
    CmdOpen.Enabled = True
    cmdPost.Caption = "&Post"
    CmdDelete.Enabled = True
    cmdPreview.Enabled = True
    cmdPost.Enabled = True
    cmdUndoPost.Enabled = True
    Call alldisable
'    If Not rsItemMaster.EOF Then FindRecord
End If
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
'Call Delete_Duplicates
End Sub


Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If frmLogin.txtUID.text = "Admin" Then
    If idelete = vbYes Then
            cn.Execute "Delete From SStockMaster Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
            cn.Execute "DELETE FROM SStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
            Call Clear
        End If
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdEdit_Click()
 
'-----------------Admin Check--------
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "select SerialNo,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & frmLogin.txtUID.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
' ----------------Check End------

If rs!Privilegegroup = 0 Then

 If txtpost.text = "Not Posted" Then
    If CmdEdit.Caption = "&Edit" Then
        CmdNew.Enabled = False
        Call allenable
        CmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdOpen.Enabled = False
'        cmdDelete.Enabled = False
        cmdPreview.Enabled = False
'        cmdLAdd.Enabled = True
        cmdRDelete.Enabled = True
        fgStock.Editable = flexEDKbdMouse
        txtSerialNo.Enabled = False
        cmdPost.Enabled = False
        
        Call Calculation
        
    ElseIf CmdEdit.Caption = "&Update" Then
          Call Calculation
          
        If IsValidRecord Then
            If rcupdate Then
                CmdEdit.Caption = "&Edit"
                CmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdOpen.Enabled = True
                cmdPreview.Enabled = True
'                cmdDelete.Enabled = True
                cmdClose.Enabled = True
                cmdPost.Enabled = True
                fgStock.Editable = flexEDNone
                Call alldisable
                
                rsItemMaster.Requery
                Dim s As String
                s = txtSerialNo
                rsItemMaster.MoveFirst
                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord
            End If
        End If
   End If
 End If

Else
' If txtpost.text = "Not Posted" Then
    If CmdEdit.Caption = "&Edit" Then
        CmdNew.Enabled = False
        Call allenable
        CmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdOpen.Enabled = False
        CmdDelete.Enabled = False
        cmdPreview.Enabled = False
'        cmdLAdd.Enabled = True
        cmdRDelete.Enabled = True
        fgStock.Editable = flexEDKbdMouse
        txtSerialNo.Enabled = False
        cmdPost.Enabled = False
        cmdUndoPost.Enabled = False
        
        Call Calculation
        
    ElseIf CmdEdit.Caption = "&Update" Then
    
    Call Calculation
'          Call duplicate
        If IsValidRecord Then
            If rcupdate Then
                CmdEdit.Caption = "&Edit"
                CmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdOpen.Enabled = True
                cmdPreview.Enabled = True
                CmdDelete.Enabled = True
                cmdClose.Enabled = True
                cmdPost.Enabled = True
                cmdUndoPost.Enabled = True
                fgStock.Editable = flexEDNone
                Call alldisable
                
                rsItemMaster.Requery
'                Dim s As String
                s = txtSerialNo
                rsItemMaster.MoveFirst
                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord
                
            End If
        End If
    End If
'  End If

End If

End Sub

Private Sub cmdLAdd_Click()
With fgStock
        If .Row = -1 Or .Row = 0 Then
            .AddItem ""
            Exit Sub
        End If
        If .Row > 0 Then
                .AddItem "", .Row + 1
        End If
    End With
    
End Sub

Private Sub cmdPost_Click()
Dim s As String
cmdPost.Caption = "&Posting"
fgStock.Editable = flexEDKbdMouse
Call postedCheck


If cmdPost.Caption = "&Posting" Then
     If txtpost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 CmdNew.Caption = "&New"
                 CmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 fgStock.Enabled = False
                 CmdOpen.Enabled = True
                 cmdPreview.Enabled = True
                 CmdDelete.Enabled = True
'                 cmdChange.Enabled = True
'                 txtSerialNo.Enabled = False
                 Call alldisable
           End If
        End If
      End If
Else
'    cmdtemSelected.Enabled = False
    cmdRDelete.Enabled = False
 End If
cmdPost.Caption = "&Post"
 
End Sub



Private Sub cmdPrevious_Click()
Dim i As Integer
    Dim strLedgerDetail As String
    Set rsItemDetail = New ADODB.Recordset

Adodc2.Recordset.MovePrevious
If Adodc2.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdPrevious.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
    txtSerialNo = rsItemMaster!SerialNo
    RDate = rsItemMaster!RDate
    cmbSName = rsItemMaster!SName
    txtSBill = rsItemMaster!SBill
    cmbDepartment = rsItemMaster!DName
    cmbStoreName = rsItemMaster!StoreName
    txtTotalBill = rsItemMaster!TotalBill
    txtPaid = rsItemMaster!Paid
    txtDue = rsItemMaster!Due
    txtpost = rsItemMaster!Posted
    txtUName = rsItemMaster!UName
        
        
    fgStock.Rows = 1
    strLedgerDetail = "SELECT  SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit " & _
                      "FROM SStockDetails " & _
                      "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsItemDetail.RecordCount <> 0 Then

        fgStock.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgStock.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgStock.TextMatrix(i, 2) = rsItemDetail("RDate")
            fgStock.TextMatrix(i, 3) = rsItemDetail("SName")
            fgStock.TextMatrix(i, 4) = rsItemDetail("StoreName")
            fgStock.TextMatrix(i, 5) = rsItemDetail("Catagory")
            fgStock.TextMatrix(i, 6) = rsItemDetail("ItemName")
            fgStock.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgStock.TextMatrix(i, 8) = rsItemDetail("Rate")
            fgStock.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgStock.TextMatrix(i, 10) = rsItemDetail("Rol")
            fgStock.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("ExpDate")), "", Format(rsItemDetail("ExpDate"), "dd-mmm-yyyy"))
            fgStock.TextMatrix(i, 12) = rsItemDetail("Posted")
            fgStock.TextMatrix(i, 13) = rsItemDetail("Warranty")
            fgStock.TextMatrix(i, 14) = rsItemDetail("Remarks")
            fgStock.TextMatrix(i, 15) = rsItemDetail("CPost")
            fgStock.TextMatrix(i, 16) = rsItemDetail("Unit")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
    End If
End Sub

Private Sub cmdUndoPost_Click()
Dim s As String
cmdUndoPost.Caption = "&Undo Posting"
fgStock.Editable = flexEDKbdMouse
Call postedCheck


If cmdUndoPost.Caption = "&Undo Posting" Then
     If txtpost.text = "Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 CmdNew.Caption = "&New"
                 CmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 fgStock.Enabled = False
                 CmdOpen.Enabled = True
                 cmdPreview.Enabled = True
                 CmdDelete.Enabled = True
'                 cmdChange.Enabled = True
'                 txtSerialNo.Enabled = False
                 Call alldisable
           End If
        End If
      End If
Else
'    cmdtemSelected.Enabled = False
    cmdRDelete.Enabled = False
 End If
cmdUndoPost.Caption = "&Undo Post"
End Sub

Private Sub fgStock_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

Select Case Col
'   Case 6, 4, 9
Case 6, 9

      Cancel = True
End Select

End Sub


Private Sub cmdRDelete_Click()
    
    
    If fgStock.Rows = 1 Then Exit Sub

     If fgStock.Row >= 1 Then
      If MsgBox("Are you sure to delete the selected record", vbYesNo, "General Setup") = vbYes Then fgStock.RemoveItem fgStock.Row
     Else
      MsgBox "You have to select a row to delete.", vbInformation, "General"
    End If
    

End Sub

Private Sub cmdNew_Click()

'-----------------Admin Check--------
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "select SerialNo,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & frmLogin.txtUID.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
' ----------------Check End------
            
   '   Dim rs As String
If rs!Privilegegroup = 0 Then

'    Set rs = New ADODB.Recordset
If CmdNew.Caption = "&New" Then
        
        CmdNew.Caption = "&Save"
        CmdEdit.Enabled = False
        cmdCancel.Enabled = True
        CmdOpen.Enabled = False
        CmdDelete.Enabled = False
        CmdOpen.Enabled = False
        cmdClose.Enabled = False
        cmdPost.Enabled = False
        cmdRDelete.Enabled = True
        cmdPreview.Enabled = False
        
        TextClear Me
        Call Clear
        RDate.Value = Date
         
        fgStock.Rows = 1
'        fgStock.Editable = flexEDKbdMouse
        Call allenable
        txtpost.text = "Not Posted"
        txtUName.text = frmLogin.txtUID.text
        cmbSName.SetFocus

        Call Calculation
    
    ElseIf CmdNew.Caption = "&Save" Then
    
    
Call Calculation
        If IsValidRecord Then
            If rcupdate Then
                CmdNew.Caption = "&New"
                CmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
'                cmdDelete.Enabled = True
                CmdOpen.Enabled = True
                cmdCancel.Enabled = True
                cmdPreview.Enabled = True
                cmdPost.Enabled = True
                
'                RDate.Value = Date
                
                Call alldisable
            End If
        End If
    End If
    
Else

' Set rs = New ADODB.Recordset
If CmdNew.Caption = "&New" Then
        
        CmdNew.Caption = "&Save"
        CmdEdit.Enabled = False
        cmdCancel.Enabled = True
        CmdOpen.Enabled = False
        CmdDelete.Enabled = False
        CmdOpen.Enabled = False
        cmdClose.Enabled = False
        cmdPost.Enabled = False
        cmdUndoPost.Enabled = False
'        cmdLAdd.Enabled = True
        cmdRDelete.Enabled = True
        cmdPreview.Enabled = False
        TextClear Me
        Call Clear
        RDate.Value = Date
         
        fgStock.Rows = 1
'        fgStock.Editable = flexEDKbdMouse
        Call allenable
        
'        RDate.Value = Date
        
        txtpost.text = "Not Posted"
        txtUName.text = frmLogin.txtUID.text
        cmbSName.SetFocus
Call Calculation
        
    ElseIf CmdNew.Caption = "&Save" Then
'        Dim rs As String
        Call Calculation
'        RDate.Value = Date
        If IsValidRecord Then
            If rcupdate Then
                CmdNew.Caption = "&New"
                CmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
                CmdOpen.Enabled = True
                cmdCancel.Enabled = True
                cmdPreview.Enabled = True
                cmdPost.Enabled = True
                cmdUndoPost.Enabled = True
                
                Call alldisable
            End If
        End If
    End If
End If
    
End Sub

Private Sub Clear()
    txtSerialNo.text = ""
'    RDate.Enabled = False
    cmbStoreName.text = ""
    cmbSName.text = ""
    txtSBill.text = ""
    txtTotalBill.text = ""
    txtPaid.text = ""
    txtDue.text = ""
    cmbDepartment.text = ""
    
End Sub

Private Sub allenable()
'     txtSerialNo.Enabled = True
     RDate.Enabled = True
     cmbSName.Enabled = True
     txtSBill.Enabled = True
     txtPaid.Enabled = True
     cmbDepartment.Enabled = True
     fgStock.Enabled = True
     cmdAddItem.Enabled = True
'     cmdLAdd.Enabled = True
     cmdRDelete.Enabled = True
    End Sub

Private Sub alldisable()
     txtSerialNo.Enabled = False
     cmdAddItem.Enabled = False
     cmdRDelete.Enabled = False
     fgStock.Enabled = False
     RDate.Enabled = False
'     RDate.Value = False
     cmbSName.Enabled = False
     txtSBill.Enabled = False
     txtTotalBill.Enabled = False
     txtPaid.Enabled = False
     txtDue.Enabled = False
     cmbDepartment.Enabled = False

    
End Sub

Private Sub cmdOpen_Click()
    frmStockSearch.Show vbModal
    Call Calculation
    CmdOpen.Enabled = True
    cmdCancel.Enabled = True
        
End Sub
    
Private Sub Command1_Click()
frmCatagory.Show vbModal
End Sub


 Private Sub Form_Load()
         Call Connect
     ModFunction.StartUpPosition Me
     txtUName.text = frmLogin.txtUID.text
       Call alldisable
       Call SName
       Call DName
       Call StoreName
       Call Calculation
       
       txtpost.text = "Not Posted"
       
   Set rsItemMaster = New ADODB.Recordset
 
  If rsItemMaster.State <> 0 Then rsItemMaster.Close
  
     rsItemMaster.Open "select SerialNo, RDate, SName, SBill, DName, StoreName, " & _
                       "TotalBill, Paid, Due, Posted, UName from SStockMaster", cn, adOpenStatic, adLockReadOnly

If rsItemMaster.RecordCount > 0 Then
      rsItemMaster.MoveLast
      
        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
    If Not rsItemMaster.EOF Then FindRecord
    
    '-----------------For Record Search----------
Adodc2.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

Adodc2.CommandType = adCmdTable
Adodc2.RecordSource = "SStockMaster"

Adodc2.Refresh
'-------------------End Record Search---------
    
    Call Calculation
    
End Sub
Private Sub FindRecord()

    Dim i As Integer
    Dim strLedgerDetail As String
    Set rsItemDetail = New ADODB.Recordset
    txtSerialNo = rsItemMaster!SerialNo
    RDate = rsItemMaster!RDate
    cmbSName = rsItemMaster!SName
    txtSBill = rsItemMaster!SBill
    cmbDepartment = rsItemMaster!DName
    cmbStoreName = rsItemMaster!StoreName
    txtTotalBill = rsItemMaster!TotalBill
    txtPaid = rsItemMaster!Paid
    txtDue = rsItemMaster!Due
    txtpost = rsItemMaster!Posted
    txtUName = rsItemMaster!UName


    fgStock.Rows = 1
    strLedgerDetail = "SELECT  SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate,Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit" & _
                " FROM SStockDetails " & _
                "WHERE SerialNo='" & parseQuotes(txtSerialNo.text) & "'"
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strLedgerDetail, cn, adOpenStatic, adLockReadOnly

 If rsItemDetail.RecordCount <> 0 Then

        fgStock.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgStock.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgStock.TextMatrix(i, 2) = rsItemDetail("RDate")
            fgStock.TextMatrix(i, 3) = rsItemDetail("SName")
            fgStock.TextMatrix(i, 4) = rsItemDetail("StoreName")
            fgStock.TextMatrix(i, 5) = rsItemDetail("Catagory")
            fgStock.TextMatrix(i, 6) = rsItemDetail("ItemName")
            fgStock.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgStock.TextMatrix(i, 8) = rsItemDetail("Rate")
            fgStock.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgStock.TextMatrix(i, 10) = rsItemDetail("Rol")
            fgStock.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("ExpDate")), "", Format(rsItemDetail("ExpDate"), "dd-mmm-yyyy"))
            fgStock.TextMatrix(i, 12) = rsItemDetail("Posted")
            fgStock.TextMatrix(i, 13) = rsItemDetail("Warranty")
            fgStock.TextMatrix(i, 14) = rsItemDetail("Remarks")
            fgStock.TextMatrix(i, 15) = rsItemDetail("CPost")
            fgStock.TextMatrix(i, 16) = rsItemDetail("Unit")
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
End Sub

Private Sub StoreName()
    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT StoreName FROM HStore ORDER BY StoreName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbStoreName.AddItem rsTemp2("StoreName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub

Private Sub DName()
    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT sDName FROM SDeptName ORDER BY sDName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbDepartment.AddItem rsTemp2("sDName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub

 Private Function rcupdate() As Boolean

'On Error GoTo ErrHandler
    Dim strSQL As String
    Dim iRow As Integer
    Dim j As Integer
    Dim i As Integer
    Dim blnAlarm As Boolean
    Dim strDeliveryDate As String
    Dim str As String
    Set rs = New ADODB.Recordset
    Dim ipost
    Dim strExpDate As String
'-------------------------------Group Permission------------
str = "select SerialNo,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & frmLogin.txtUID.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
'           If rs.RecordCount = 0 Then Exit Sub
'-------------------------------Group permission end-------------------
    If rs!Privilegegroup = 0 Then
     
'     cn.BeginTrans
     If CmdNew.Caption = "&Save" Then
        
    'General Information for Payment Master
    
     strSQL = "INSERT INTO SStockMaster (RDate, SName, SBill, DName, StoreName, TotalBill, Paid, Due, Posted, UName " & _
                ") " & _
                "VALUES ('" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "', " & _
                " '" & txtSBill & "','" & cmbDepartment.text & "','" & cmbStoreName.text & "'," & Val(txtTotalBill.text) & ", " & _
                " " & Val(txtPaid.text) & "," & Val(txtDue.text) & ",'" & txtpost & "','" & txtUName.text & "')"
     cn.Execute strSQL
    

      rcupdate = True
'     cn.CommitTrans
     
'     -------------For primary key and foreign key relation------------
         If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),1) as InvNo from SStockMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo = Val(rs!InvNo)


            j = 0
            For j = 1 To fgStock.Rows - 1
            
            If fgStock.Cell(flexcpChecked, j, 12) = flexChecked Then
               blnAlarm = True
            Else
                blnAlarm = False
            End If
            
    cn.Execute "INSERT INTO SStockDetails (SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, " & _
               "Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit) " & _
               "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
               "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
               IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & ", " & _
               IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
               IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & ", " & _
               IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
               IIf(fgStock.TextMatrix(j, 11) = "", "NUll", "'" & fgStock.TextMatrix(j, 11) & "' ") & ", " & _
               IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 13)) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 16)) & "')"
            Next
        
        rcupdate = True
        
        
        
'        cn.CommitTrans
        MsgBox "Record added Successfully", vbInformation, "Confirmation"
    
    ' Update Information
    
    

ElseIf (CmdEdit.Caption = "&Update") Then
    
'            If txtpost.text = "Not Posted" Then
            
    cn.Execute "UPDATE SStockMaster SET  RDate = '" & Format(RDate, "dd-mmm-yyyy") & "', " & _
               "SName='" & cmbSName.text & "',SBill='" & txtSBill & "',DName='" & cmbDepartment.text & "',StoreName='" & cmbStoreName.text & "', " & _
               "TotalBill=" & Val(txtTotalBill.text) & ",Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ", " & _
               "Posted='" & txtpost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM SStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


        j = 0
            For j = 1 To fgStock.Rows - 1
            
            If fgStock.Cell(flexcpChecked, j, 12) = flexChecked Then
               blnAlarm = True
            Else
                blnAlarm = False
            End If
             
    cn.Execute "INSERT INTO SStockDetails (SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, " & _
               "Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit) " & _
               "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
               "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
               IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & ", " & _
               IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
               IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & ", " & _
               IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
               IIf(fgStock.TextMatrix(j, 11) = "", "NUll", "'" & fgStock.TextMatrix(j, 11) & "' ") & ", " & _
               IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 13)) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 16)) & "')"
               Next

        rcupdate = True
'        cn.CommitTrans
        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
        
'        End If
        
        
'  --------------------------------Posting Information-----------------------------------------

   ElseIf cmdPost.Caption = "&Posting" Then
   
     
'''     Dim iPost
     txtpost.text = "Posted"
     
     

ipost = MsgBox("Do you want to Post this bill?", vbYesNo)

If ipost = vbYes Then
     
     txtpost.text = "Posted"
     
    cn.Execute "UPDATE SStockMaster SET  RDate = '" & Format(RDate, "dd-mmm-yyyy") & "', " & _
                "SName='" & cmbSName.text & "',SBill='" & txtSBill & "',DName='" & cmbDepartment.text & "',StoreName='" & cmbStoreName.text & "', " & _
                "TotalBill=" & Val(txtTotalBill.text) & ",Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ", " & _
                "Posted='" & txtpost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM SStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


       j = 0
            For j = 1 To fgStock.Rows - 1
            
            If fgStock.Cell(flexcpChecked, j, 12) = flexChecked Then
               blnAlarm = True
            Else
                blnAlarm = False
            End If
             
  cn.Execute "INSERT INTO SStockDetails (SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, " & _
               "Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit) " & _
               "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
               "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
               IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & ", " & _
               IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
               IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & ", " & _
               IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
               IIf(fgStock.TextMatrix(j, 11) = "", "NUll", "'" & fgStock.TextMatrix(j, 11) & "' ") & ", " & _
               IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 13)) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 16)) & "')"
               Next

        rcupdate = True
'        cn.CommitTrans
        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
        End If
    
    End If

Else

'-------------Admin group------------------

  cn.BeginTrans
     If CmdNew.Caption = "&Save" Then
        
  strSQL = "INSERT INTO SStockMaster (RDate, SName, SBill, DName, StoreName, TotalBill, Paid, Due, Posted, UName " & _
            ") " & _
            "VALUES ('" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "', " & _
            " '" & txtSBill & "','" & cmbDepartment.text & "','" & cmbStoreName.text & "'," & Val(txtTotalBill.text) & ", " & _
            " " & Val(txtPaid.text) & "," & Val(txtDue.text) & ",'" & txtpost & "','" & txtUName.text & "')"
     
     cn.Execute strSQL
      rcupdate = True
'     cn.CommitTrans
     
'     -------------For primary key and foreign key relation------------
         If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),1) as InvNo from SStockMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo = Val(rs!InvNo)

            j = 0
            For j = 1 To fgStock.Rows - 1
            
            If fgStock.Cell(flexcpChecked, j, 12) = flexChecked Then
               blnAlarm = True
            Else
                blnAlarm = False
            End If
             If fgStock.TextMatrix(i, 11) = "" Then
            strExpDate = "null"
            End If


   cn.Execute "INSERT INTO SStockDetails (SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, " & _
               "Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit) " & _
               "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
               "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
               IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & ", " & _
               IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
               IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & ", " & _
               IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
               IIf(fgStock.TextMatrix(j, 11) = "", "NUll", "'" & fgStock.TextMatrix(j, 11) & "' ") & ", " & _
               IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 13)) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 16)) & "')"
               
               Next
        rcupdate = True
        
        cn.CommitTrans
        MsgBox "Record added Successfully", vbInformation, "Confirmation"
    
ElseIf (CmdEdit.Caption = "&Update") Then
    
cn.Execute "UPDATE SStockMaster SET  RDate = '" & Format(RDate, "dd-mmm-yyyy") & "', " & _
            "SName='" & cmbSName.text & "',SBill='" & txtSBill & "',DName='" & cmbDepartment.text & "',StoreName='" & cmbStoreName.text & "', " & _
            "TotalBill=" & Val(txtTotalBill.text) & ",Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ", " & _
            "Posted='" & txtpost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM SStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


        j = 0
            For j = 1 To fgStock.Rows - 1
            
            If fgStock.Cell(flexcpChecked, j, 12) = flexChecked Then
               blnAlarm = True
            Else
                blnAlarm = False
            End If
            
 cn.Execute "INSERT INTO SStockDetails (SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, " & _
            "Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit) " & _
            "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
            "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
            IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & ", " & _
            IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
            IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & ", " & _
            IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
            IIf(fgStock.TextMatrix(j, 11) = "", "NUll", "'" & fgStock.TextMatrix(j, 11) & "' ") & ", " & _
            IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 13)) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 16)) & "')"
               Next

        rcupdate = True
        cn.CommitTrans
        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
        
'        End If
        
'  --------------------------------Posting Information-----------------------------------------


   ElseIf cmdPost.Caption = "&Posting" Then
   
     
'''     Dim iPost
     txtpost.text = "Posted"
     
     

ipost = MsgBox("Do you want to Post this bill?", vbYesNo)

If ipost = vbYes Then
     
     txtpost.text = "Posted"
     
  cn.Execute "UPDATE SStockMaster SET  RDate = '" & Format(RDate, "dd-mmm-yyyy") & "', " & _
               "SName='" & cmbSName.text & "',SBill='" & txtSBill & "',DName='" & cmbDepartment.text & "',StoreName='" & cmbStoreName.text & "', " & _
               "TotalBill=" & Val(txtTotalBill.text) & ",Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ", " & _
               "Posted='" & txtpost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM SStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


       j = 0
            For j = 1 To fgStock.Rows - 1
            
            If fgStock.Cell(flexcpChecked, j, 12) = flexChecked Then
               blnAlarm = True
            Else
                blnAlarm = False
            End If
    
cn.Execute "INSERT INTO SStockDetails (SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, " & _
             "Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit) " & _
             "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
             "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
             IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & ", " & _
             IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
             IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & ", " & _
             IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
             IIf(fgStock.TextMatrix(j, 11) = "", "NUll", "'" & fgStock.TextMatrix(j, 11) & "' ") & ", " & _
             IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 13)) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 16)) & "')"
               Next

        rcupdate = True
        cn.CommitTrans
        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
        End If
    
    
    
    
'  -----------Undo Posting-------

ElseIf cmdUndoPost.Caption = "&Undo Posting" Then
   
     
'''     Dim iPost
     txtpost.text = "Not Posted"
     
     

ipost = MsgBox("Do you want to Undo post this bill?", vbYesNo)

If ipost = vbYes Then
     
     txtpost.text = "Not Posted"
      cn.Execute "UPDATE SStockMaster SET  RDate = '" & Format(RDate, "dd-mmm-yyyy") & "', " & _
                "SName='" & cmbSName.text & "',SBill='" & txtSBill & "',DName='" & cmbDepartment.text & "',SubCode='" & cmbStoreName.text & "', " & _
                "TotalBill=" & Val(txtTotalBill.text) & ",Paid=" & Val(txtPaid.text) & ",Due=" & Val(txtDue.text) & ", " & _
                "Posted='" & txtpost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM SStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


       j = 0
            For j = 1 To fgStock.Rows - 1
            
            If fgStock.Cell(flexcpChecked, j, 12) = flexChecked Then
               blnAlarm = True
            Else
                blnAlarm = False
            End If
 
 
    cn.Execute "INSERT INTO SStockDetails (SerialNo, RDate, SName, StoreName, Catagory, ItemName, Qty, Rate, " & _
               "Amount, Rol, ExpDate, Posted, Warranty, Remarks, CPost, Unit) " & _
               "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(RDate, "dd-mmm-yyyy") & "','" & cmbSName.text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
               "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
               IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & ", " & _
               IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
               IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & ", " & _
               IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
               IIf(fgStock.TextMatrix(j, 11) = "", "NUll", "'" & fgStock.TextMatrix(j, 11) & "' ") & ", " & _
               IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 13)) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 16)) & "')"
               Next

        rcupdate = True
        cn.CommitTrans
        MsgBox "Record Undo Posted Successfully", vbInformation, "Confirmation"
        
        End If


'-------------Undo Posting End----
    
    End If

'-------------Admin group end--------------


End If

'    cn.CommitTrans
    
    Exit Function
    
ErrHandler:

    cn.RollbackTrans
    Select Case Err.Number
        Case -2147217900
            MsgBox "Please select Numeric number in ROL field", vbInformation, "Confirmation"

   End Select
   

End Function

Private Sub Calculation()

  Dim j As Integer
  Dim i As Integer
 
       temp = 0
'       temp1 = 0
'       temp2 = 0
    For j = 1 To fgStock.Rows - 1

        temp = temp + CDbl(Val(fgStock.TextMatrix(j, 7)) * CDbl(Val(fgStock.TextMatrix(j, 8))))
   Next
   
'   txtTotalBill = sum(Val(fgStock.TextMatrix(j, 9)))
   txtTotalBill = temp


txtDue = (CDbl(txtTotalBill) - CDbl(Val(txtPaid)))
'txtWords = InWords(txtNPayable.text)

End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    If Trim(RDate) = "" Then
        MsgBox "Your are missing Receiving Information.", vbInformation
        RDate.SetFocus
        IsValidRecord = False
        Exit Function
 
        
  ElseIf Trim(cmbSName) = "" Then
        MsgBox "Your are missing Supplier Name Information.", vbInformation
        cmbSName.SetFocus
        IsValidRecord = False
        Exit Function
        
        
 ElseIf Trim(txtSBill) = "" Then
        MsgBox "Your are missing Supplier Bill No.", vbInformation
        txtSBill.SetFocus
        IsValidRecord = False
        Exit Function
        
'-----------------------------------------------------------------------
    Else
        
       
       Exit Function
     End If
    End Function
    



Public Sub printReport()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    If rsItemMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\ReceivingReceipt.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        


    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close


                      
                      
strSQL = "SELECT SStockMaster.SerialNo, SStockMaster.RDate,SStockMaster.SName,SStockMaster.SBill, " & _
          "SStockMaster.DName, SStockDetails.StoreName, SStockDetails.Catagory, SStockDetails.ItemName, " & _
          "SStockDetails.Qty,SStockDetails.Rate, SStockDetails.Amount, SStockDetails.ExpDate, " & _
          "SStockDetails.Posted , SStockDetails.Warranty, SStockDetails.Remarks,SStockMaster.UName " & _
          "FROM SStockMaster INNER JOIN " & _
          "SStockDetails ON SStockMaster.SerialNo = SStockDetails.SerialNo and SStockMaster.SerialNo ='" & Me.txtSerialNo & "' ORDER BY SStockDetails.Catagory "

                      rsDailyRpt.Open strSQL, cn, adOpenStatic
'        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"


        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Menu Item List Report", , , , , 16777216 Or 524288 Or 65536
    
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

ErrH:

    Select Case Err.Number
        Case -2147217913
            MsgBox "You need to select record first", vbInformation, "Item Catagory List Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Item catagory Report"
    End Select
End Sub


Private Sub duplicate()
   Dim j As Integer
        
         For j = 1 To fgStock.Rows - 2
        
        If Val(fgStock.TextMatrix(j, 4)) = Val(fgStock.TextMatrix(j + 1, 4)) Then
        MsgBox "Duplicate Item Catagory Name.", vbInformation
         fgStock.TextMatrix(j, 4) = ""
         End If

         Next

End Sub

Public Sub PopulateForm(StrID As String)
    rsItemMaster.Close
    rsItemMaster.Open "select * from SStockMaster", cn, adOpenStatic, adLockReadOnly
    rsItemMaster.MoveFirst
    rsItemMaster.Find "SerialNo=" & parseQuotes(StrID)
    If rsItemMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


Private Sub fgStock_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim j As Integer
     Select Case Col
        Case 9
'------------------duplicate folio-----------
Case 6
    Dim k As Integer
If fgStock.Rows > 2 Then
    For k = 1 To fgStock.Rows - 1
 If (fgStock.TextMatrix(k, 6)) = fgStock.TextMatrix(Row, 6) And k <> fgStock.Row Then
        MsgBox "Duplicate Item Name.", vbInformation
        fgStock.TextMatrix(Row, 6) = ""

     End If

   Next
End If
'-------------------------------------------
End Select

Call Calculation
End Sub

Private Sub txtPaid_Change()
'If KeyAscii = 13 Then
'SendKeys Chr(9)
'End If
Call Calculation
End Sub

Private Sub txtSBill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub


Private Sub txtSBill_GotFocus()
txtSBill.BackColor = &HFFFFC0
End Sub

Private Sub txtSBill_LostFocus()
    txtSBill.BackColor = vbWhite
End Sub
