VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRequisition 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Requisitions"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14160
   Icon            =   "frmRequisition.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUndoPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Undo Post"
      Height          =   735
      Left            =   11040
      Picture         =   "frmRequisition.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9960
      Width           =   1095
   End
   Begin VB.CommandButton CmdPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Post"
      Height          =   735
      Left            =   10080
      Picture         =   "frmRequisition.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   735
      Left            =   9120
      Picture         =   "frmRequisition.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton chameleonButton1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      Height          =   735
      Left            =   8160
      Picture         =   "frmRequisition.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Print"
      Height          =   735
      Left            =   7080
      Picture         =   "frmRequisition.frx":2334
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   9960
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Q&uit"
      Height          =   735
      Left            =   6120
      Picture         =   "frmRequisition.frx":2BFE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   735
      Left            =   5160
      Picture         =   "frmRequisition.frx":34C8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      Height          =   735
      Left            =   4200
      Picture         =   "frmRequisition.frx":3D92
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   735
      Left            =   3240
      Picture         =   "frmRequisition.frx":4432
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   735
      Left            =   2280
      Picture         =   "frmRequisition.frx":4CFC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9960
      Width           =   990
   End
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
      Left            =   14640
      Picture         =   "frmRequisition.frx":55C6
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2520
      Width           =   420
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Requisation Master Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   15015
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
         Left            =   8160
         Picture         =   "frmRequisition.frx":5B50
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtSerialNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   1920
         TabIndex        =   14
         Text            =   " "
         Top             =   360
         Width           =   3135
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   2535
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbBudgetHead 
         Height          =   405
         Left            =   9120
         TabIndex        =   16
         Top             =   240
         Width           =   3105
         AllowInput      =   0   'False
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5477
         _ExtentY        =   714
         _StockProps     =   93
         BackColor       =   -2147483643
      End
      Begin MSComCtl2.DTPicker ReceivingDate 
         Height          =   405
         Left            =   1920
         TabIndex        =   17
         Top             =   1080
         Width           =   3105
         _ExtentX        =   5477
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
         Format          =   65536003
         CurrentDate     =   37840
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Requisation Date"
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
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lblBudgetHead 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Department Name"
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
         Left            =   7440
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblReqNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Requisation No"
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
         TabIndex        =   18
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Requisation Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   7455
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   15015
      Begin VSFlex7DAOCtl.VSFlexGrid fgStock 
         Height          =   6975
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   14535
         _cx             =   25638
         _cy             =   12303
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
         BackColorFixed  =   -2147483630
         ForeColorFixed  =   49152
         BackColorSel    =   -2147483630
         ForeColorSel    =   -2147483634
         BackColorBkg    =   12629161
         BackColorAlternate=   -2147483629
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
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRequisition.frx":641A
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.TextBox txtUserName 
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
      Height          =   525
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   9960
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc dcBudgetHead 
      Height          =   330
      Left            =   120
      Top             =   9840
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
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
      Caption         =   "dcBudgetHead"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   600
      Left            =   0
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
   Begin MSAdodcLib.Adodc dcSupplierName 
      Height          =   330
      Left            =   120
      Top             =   10200
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
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
      Caption         =   "dcSupplierName"
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      Caption         =   "Store Management System (SMS) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
' Private rsItemMaster                 As ADODB.Recordset
' Private rsItemDetail                 As ADODB.Recordset
' Private rs                              As ADODB.Recordset
' Private bRecordExists                  As Boolean
' Dim str As String
''---------------------------------------------------------------------------
''---------------------------------------------------------------------------
''----Add For Reporting Perpose----------------------------------------------
'Private objReportApp                        As CRPEAuto.Application
'Private objReport                           As CRPEAuto.Report
'Private objReportDatabase                   As CRPEAuto.Database
'Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
'Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
'Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private objReportFF                         As CRPEAuto.FormulaFieldDefinition
'
'
'Private objReportSub                        As CRPEAuto.Report 'sub
'Private objReportDatabaseSub                As CRPEAuto.Database 'sub
'Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
'Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
'Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
'Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition
'
'
'Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
'Private rsDailyRpt                          As ADODB.Recordset
'Private Tracer                              As Integer
'Private strGroupName                        As String
'Private temp As Double
'Private temp1 As Double
''--------------------------------------------------------------------------------
'
'
'Private Sub chameleonButton1_Click()
'    Call printReport
'End Sub
'
''Private Sub Check1_Click()
''Dim iRows As Integer
''Dim f As Integer
''      f = fgStock.Rows - 1
''  If Check1.Value = 1 Then
''       Dim i As Integer
''    For i = 1 To f
''      fgStock.Cell(flexcpChecked, i, 10) = flexChecked
''
''      Next i
''Else
''     For i = 1 To f
''
''        fgStock.Cell(flexcpChecked, i, 10) = flexUnchecked
''
''    Next i
''    End If
''End Sub
'
''Private Sub chkAutoposting_Click()
''
''      Dim f As Integer
''      f = fgStock.Rows - 1
''      If chkAutoposting.Value = 1 Then
''      Dim i As Integer
''For i = 1 To f
''    fgStock.Cell(flexcpChecked, i, 12) = flexChecked
''Next i
''Else
''For i = 1 To f
''    fgStock.Cell(flexcpChecked, i, 12) = flexUnchecked
''
''    Next i
''End If
''
''End Sub
'
'Private Sub cmdAddItem_Click()
''ModFunction.StartUpPosition Me
'frmItemReqReceiving.Show vbModal
''frmItemReqReceiving.Show
'End Sub
'
'
'Private Sub postedCheck()
'      Dim f As Integer
'      Dim i As Integer
'      f = fgStock.Rows - 1
''      If chkAutoposting.Value = 1 Then
'
'For i = 1 To f
'    fgStock.Cell(flexcpChecked, i, 11) = flexChecked
'Next i
'
'
'End Sub
'Private Sub cmdCancel_Click()
'
'    cmdCancel.Enabled = False
'    CmdNew.Enabled = True
'    CmdEdit.Caption = "&Edit"
'    CmdNew.Caption = "&New"
'    cmdClose.Enabled = True
'    CmdEdit.Enabled = True
'    cmdOpen.Enabled = True
'    cmdPost.Caption = "&Post"
''    cmdDelete.Enabled = True
'    chameleonButton1.Enabled = True
'    Call alldisable
''    If Not rsItemMaster.EOF Then FindRecord
'
'End Sub
'
'
'Private Sub cmdClose_Click()
'    Unload Me
''Call Delete_Duplicates
'End Sub
'
'
'Private Sub cmdDelete_Click()
'On Error GoTo ErrHandler
'     Dim idelete As Integer
'    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
'    If idelete = vbYes Then
'            cn.Execute "Delete From SReqMaster Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
'            cn.Execute "DELETE FROM SReqDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'            Call Clear
'    End If
'ErrHandler:
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
'     End Select
'End Sub
'
'Private Sub cmdEdit_Click()
''-----------------Admin Check--------
'Dim s As String
'Set rs = New ADODB.Recordset
'If rs.State <> 0 Then rs.Close
'str = "select SerialNo,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & frmLogin.txtUID.text & "'"
'         If rs.State <> 0 Then rs.Close
'            rs.Open str, cn, adOpenStatic, adLockReadOnly
'' ----------------Check End------
'
'If rs!Privilegegroup = 0 Then
'
' If txtpost.text = "Not Posted" Then
'    If CmdEdit.Caption = "&Edit" Then
'        CmdNew.Enabled = False
'        Call allenable
'        CmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        cmdOpen.Enabled = False
''        cmdDelete.Enabled = False
'        chameleonButton1.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdRDelete.Enabled = True
'        fgStock.Editable = flexEDKbdMouse
'        txtSerialNo.Enabled = False
'        cmdPost.Enabled = False
'
'    ElseIf CmdEdit.Caption = "&Update" Then
''          Call duplicate
'        If IsValidRecord Then
'            If rcupdate Then
'                CmdEdit.Caption = "&Edit"
'                CmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdOpen.Enabled = True
'                chameleonButton1.Enabled = True
''                cmdDelete.Enabled = True
'                cmdClose.Enabled = True
'                cmdPost.Enabled = True
'                fgStock.Editable = flexEDNone
'                Call alldisable
'
'                rsItemMaster.Requery
''                Dim s As String
'                s = txtSerialNo
'                rsItemMaster.MoveFirst
'                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
'                FindRecord
'            End If
'        End If
'   End If
' End If
'
'Else
'' If txtpost.text = "Not Posted" Then
'    If CmdEdit.Caption = "&Edit" Then
'        CmdNew.Enabled = False
'        Call allenable
'        CmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        cmdOpen.Enabled = False
'        CmdDelete.Enabled = False
'        chameleonButton1.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdRDelete.Enabled = True
'        fgStock.Editable = flexEDKbdMouse
'        txtSerialNo.Enabled = False
'        cmdPost.Enabled = False
'        cmdUndoPost.Enabled = False
'
'    ElseIf CmdEdit.Caption = "&Update" Then
''          Call duplicate
'        If IsValidRecord Then
'            If rcupdate Then
'                CmdEdit.Caption = "&Edit"
'                CmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdOpen.Enabled = True
'                chameleonButton1.Enabled = True
'                CmdDelete.Enabled = True
'                cmdClose.Enabled = True
'                cmdPost.Enabled = True
'                cmdUndoPost.Enabled = True
'                fgStock.Editable = flexEDNone
'                Call alldisable
'
'                rsItemMaster.Requery
''                Dim s As String
'                s = txtSerialNo
'                rsItemMaster.MoveFirst
'                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
'                FindRecord
'
'            End If
'        End If
'    End If
''  End If
'
'End If
'
'
''----------Disablale------
'' If txtpost.text = "Not Posted" Then
''    If cmdEdit.Caption = "&Edit" Then
''        cmdNew.Enabled = False
''        Call allenable
''        cmdEdit.Caption = "&Update"
''        cmdCancel.Enabled = True
''        cmdClose.Enabled = False
''        cmdOpen.Enabled = False
'''        cmdDelete.Enabled = False
''        chameleonButton1.Enabled = False
'''        cmdLAdd.Enabled = True
''        cmdLDelete.Enabled = True
''        fgStock.Editable = flexEDKbdMouse
''        txtSerialNo.Enabled = False
''
''    ElseIf cmdEdit.Caption = "&Update" Then
'''          Call duplicate
''        If IsValidRecord Then
''            If rcupdate Then
''                cmdEdit.Caption = "&Edit"
''                cmdNew.Enabled = True
''                cmdCancel.Enabled = False
''                cmdClose.Enabled = True
''                cmdOpen.Enabled = True
''                chameleonButton1.Enabled = True
'''                cmdDelete.Enabled = True
''                cmdClose.Enabled = True
''                fgStock.Editable = flexEDNone
''                Call alldisable
''            End If
''        End If
''   End If
''End If
'
'End Sub
'
'Private Sub cmdLAdd_Click()
'With fgStock
'        If .Row = -1 Or .Row = 0 Then
'            .AddItem ""
'            Exit Sub
'        End If
'        If .Row > 0 Then
'                .AddItem "", .Row + 1
'        End If
'    End With
'
'End Sub
'
'Private Sub cmdPost_Click()
'Dim s As String
'cmdPost.Caption = "&Posting"
'fgStock.Editable = flexEDKbdMouse
'
'
'Call postedCheck
'
'
'If cmdPost.Caption = "&Posting" Then
'     If txtpost.text = "Not Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 CmdNew.Caption = "&New"
'                 CmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 fgStock.Enabled = False
'                 cmdOpen.Enabled = True
'                 chameleonButton1.Enabled = True
'                 CmdDelete.Enabled = True
''                 cmdChange.Enabled = True
''                 txtBillSerialNo.Enabled = False
'                 Call alldisable
'           End If
'        End If
'      End If
'Else
''    cmdtemSelected.Enabled = False
'    cmdRDelete.Enabled = False
' End If
'cmdPost.Caption = "&Post"
'
'End Sub
'
'Private Sub cmdUndoPost_Click()
'Dim s As String
'cmdUndoPost.Caption = "&Undo Posting"
'fgStock.Editable = flexEDKbdMouse
'Call postedCheck
'
'
'If cmdUndoPost.Caption = "&Undo Posting" Then
'     If txtpost.text = "Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 CmdNew.Caption = "&New"
'                 CmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 fgStock.Enabled = False
'                 cmdOpen.Enabled = True
'                 chameleonButton1.Enabled = True
'                 CmdDelete.Enabled = True
''                 cmdChange.Enabled = True
''                 txtBillSerialNo.Enabled = False
'                 Call alldisable
'           End If
'        End If
'      End If
'Else
''    cmdtemSelected.Enabled = False
'    cmdRDelete.Enabled = False
' End If
'cmdUndoPost.Caption = "&Undo Post"
'
'End Sub
'
'Private Sub fgStock_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'   Dim pt As POINTAPI
'
'    ' get popup window position
'    pt.x = fgStock.ColPos(Col) \ Screen.TwipsPerPixelX
'    pt.y = (fgStock.RowPos(Row) + fgStock.RowHeight(Row)) \ Screen.TwipsPerPixelY
'    ClientToScreen fgStock.hwnd, pt
'
'    ' show date popup
'    If fgStock.ColDataType(Col) = flexDTDate Then
''      If Col = 9 Then
'        With frmDate
'            .lblRow = Row
'            .lblCol = Col
'            Set rsServerDate = New ADODB.Recordset
'            rsServerDate.Open "select getdate()", cn, adOpenStatic, adLockReadOnly
'            rsServerDate.Requery
'            .Tag = IIf(fgStock.Cell(flexcpText, Row, Col) = "", rsServerDate(0), fgStock.Cell(flexcpText, Row, Col))
'            strCallingForm = LCase("frmStock")
'            .Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
'            .Show vbModal
'        End With
'        Exit Sub
''       End If
'    End If
'End Sub
'
'Private Sub cmdLDelete_Click()
'
'
'    If fgStock.Rows = 1 Then Exit Sub
'
'     If fgStock.Row >= 1 Then
'      If MsgBox("Are you sure to delete the selected record", vbYesNo, "General Setup") = vbYes Then fgStock.RemoveItem fgStock.Row
'     Else
'      MsgBox "You have to select a row to delete.", vbInformation, "General"
'    End If
'
'
'End Sub
'
'Private Sub cmdNew_Click()
'
'    Set rs = New ADODB.Recordset
'If CmdNew.Caption = "&New" Then
'
'        CmdNew.Caption = "&Save"
'        CmdEdit.Enabled = False
'        cmdCancel.Enabled = True
'        cmdOpen.Enabled = False
''        cmdDelete.Enabled = False
'        cmdOpen.Enabled = False
'        cmdClose.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdRDelete.Enabled = True
'        chameleonButton1.Enabled = False
'        TextClear Me
'        Call Clear
'
'        fgStock.Rows = 1
'        fgStock.Editable = flexEDKbdMouse
'        Call allenable
'        txtpost.text = "Not Posted"
'        txtUserName.text = frmLogin.txtUID.text
''        cmbItemCatagory.SetFocus
'
'
'    ElseIf CmdNew.Caption = "&Save" Then
'        Dim s As String
''        Call duplicate1
'        If IsValidRecord Then
'            If rcupdate Then
'                CmdNew.Caption = "&New"
'                CmdEdit.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
''                cmdDelete.Enabled = True
'                cmdOpen.Enabled = True
'                cmdCancel.Enabled = True
'                chameleonButton1.Enabled = True
'
'                Call alldisable
'            End If
'        End If
'    End If
'
'End Sub
'
'Private Sub Clear()
'    txtSerialNo.text = ""
'    ReceivingDate.Enabled = False
''    cmbSupplierName.text = ""
''    txtSupplierBill.text = ""
'    cmbBudgetHead.text = ""
'
'End Sub
'
'Private Sub allenable()
''     txtSerialNo.Enabled = True
'     ReceivingDate.Enabled = True
'     ReceivingDate.Value = Date
''     cmbSupplierName.Enabled = True
''     txtSupplierBill.Enabled = True
'     cmbBudgetHead.Enabled = True
'     fgStock.Enabled = True
'     cmdAddItem.Enabled = True
''     cmdLAdd.Enabled = True
'     cmdRDelete.Enabled = True
'    End Sub
'
'Private Sub alldisable()
'     txtSerialNo.Enabled = False
''    cmbItemCatagory.Enabled = False
''     cmdLAdd.Enabled = False
'     cmdAddItem.Enabled = False
'     cmdRDelete.Enabled = False
'     fgStock.Enabled = False
'     ReceivingDate.Enabled = False
'     ReceivingDate.Value = Date
'     cmbBudgetHead.Enabled = False
'
'
'End Sub
'
'Private Sub cmdOpen_Click()
'    frmRequisitionearch.Show vbModal
'    cmdOpen.Enabled = True
'    cmdCancel.Enabled = True
'
'End Sub
'
'Private Sub Command1_Click()
'frmCatagory.Show vbModal
'End Sub
'
'
' Private Sub Form_Load()
'         Call Connect
'     ModFunction.StartUpPosition Me
'     txtUserName.text = frmLogin.txtUID.text
'       Call alldisable
''       Call SupplierName
'       Call BudgetName
'   Set rsItemMaster = New ADODB.Recordset
'
'  If rsItemMaster.State <> 0 Then rsItemMaster.Close
'     rsItemMaster.Open "select * FROM SSubCatagoryMaster", cn, adOpenStatic, adLockReadOnly
'
'If rsItemMaster.RecordCount > 0 Then
'      rsItemMaster.MoveFirst
'        bRecordExists = True
'    Else
'        bRecordExists = False
'    End If
'    txtpost.text = "Not Posted"
'
''     fgStock.ColDataType(10) = flexDTBoolean
''    If Not rsItemMaster.EOF Then FindRecord
''    txtSerialNo.Enabled = False
''    ReceivingDate.Value = Null
'End Sub
'
'
''Private Sub SupplierName()
''    dcSupplierName.CursorLocation = adUseClient
''    dcSupplierName.ConnectionString = cn.ConnectionString
''    dcSupplierName.LockType = adLockReadOnly
''    dcSupplierName.RecordSource = "SELECT sName as SupplierName , sAddress as Address FROM SMSSuplierName ORDER BY Address"
''    cmbSupplierName.DataMode = ssDataModeBound
''    Set cmbSupplierName.DataSource = dcSupplierName
''    cmbSupplierName.DataSourceList = dcSupplierName
''    cmbSupplierName.DataFieldList = "SupplierName"
'''    cmbSupplierName.DataField = "CatagoryName"
''    cmbSupplierName.BackColorOdd = &HFFFF00
''    cmbSupplierName.BackColorEven = &HFFC0C0
''    cmbSupplierName.ForeColorEven = &H80000008
''End Sub
'
'Private Sub BudgetName()
'    dcBudgetHead.CursorLocation = adUseClient
'    dcBudgetHead.ConnectionString = cn.ConnectionString
'    dcBudgetHead.LockType = adLockReadOnly
'    dcBudgetHead.RecordSource = "SELECT sDName as DepartmentName FROM SDeptName ORDER BY DepartmentName"
'    cmbBudgetHead.DataMode = ssDataModeBound
'    Set cmbBudgetHead.DataSource = dcBudgetHead
'    cmbBudgetHead.DataSourceList = dcBudgetHead
'    cmbBudgetHead.DataFieldList = "DepartmentName"
'    cmbBudgetHead.DataField = "DepartmentName"
'    cmbBudgetHead.BackColorOdd = &HFFFF00
'    cmbBudgetHead.BackColorEven = &HFFC0C0
'    cmbBudgetHead.ForeColorEven = &H80000008
'End Sub
'
' Private Function rcupdate() As Boolean
'
''On Error GoTo ErrHandler
'    Dim strSQL As String
'    Dim iRow As Integer
'    Dim j As Integer
'    Dim i As Integer
'    Dim blnAlarm As Boolean
'    Dim strDeliveryDate As String
'     Set rs = New ADODB.Recordset
'    Dim ipost
'
'
''-------------------------------Group Permission------------
'str = "select SerialNo,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & frmLogin.txtUID.text & "'"
'         If rs.State <> 0 Then rs.Close
'            rs.Open str, cn, adOpenStatic, adLockReadOnly
''           If rs.RecordCount = 0 Then Exit Sub
'
'
''--------------------------------------------------------------Group permission end-------------------
'    If rs!Privilegegroup = 0 Then
'
'   cn.BeginTrans
''     flagSlNo = 0
'    If CmdNew.Caption = "&Save" Then
'
'    'General Information for Payment Master
'     strSQL = "INSERT INTO SReqMaster (ReqDate,DeptName, CPost,UserName " & _
'                ") " & _
'                "VALUES ('" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                " '" & cmbBudgetHead.Columns(0).text & "','" & txtpost & "','" & txtUserName.text & "')"
'     cn.Execute strSQL
'      rcupdate = True
''     cn.CommitTrans
'
''     -------------For primary key and foreign key relation------------
'         If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(SerialNo),1) as InvNo from SReqMaster"
'           rs.Open str, cn, adOpenStatic, adLockReadOnly
'           txtSerialNo = Val(rs!InvNo)
'
''------------------------
'
''     '" & parseQuotes(txtSerialNo) & "',
'    'payment Detail Information Enter This table
''    strDeliveryDate = "'" & parseQuotes(fgStock.TextMatrix(j, 11)) & "'"
'            j = 0
'            For j = 1 To fgStock.Rows - 1
'
'            If fgStock.Cell(flexcpChecked, j, 11) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'''                If fgExport.TextMatrix(j, 1) <> "" Then
'                cn.Execute "INSERT INTO SReqDetails (SerialNo,ReqDate,DeptName,Catagory,SubCatagory,ItemName,Qty,Rate, " & _
'                           "Amount,Rol,Posted,Remarks,CPosted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbBudgetHead.Columns(0).text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & "," & IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & "," & IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 12)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "')"
'               Next
'        rcupdate = True
'
'
'
'        cn.CommitTrans
'        MsgBox "Record added Successfully", vbInformation, "Confirmation"
'
'    ' Update Information
'
'
'
'ElseIf (CmdEdit.Caption = "&Update") Then
'
''            If txtpost.text = "Not Posted" Then
'
'                   cn.Execute "UPDATE SReqMaster SET  ReqDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              " " & _
'                              "DeptName='" & cmbBudgetHead.Columns(0).text & "',CPost='" & txtpost & "', UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SReqDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'         j = 0
'            For j = 1 To fgStock.Rows - 1
'
'            If fgStock.Cell(flexcpChecked, j, 11) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'''                If fgExport.TextMatrix(j, 1) <> "" Then
'                cn.Execute "INSERT INTO SReqDetails (SerialNo,ReqDate,DeptName,Catagory,SubCatagory,ItemName,Qty,Rate, " & _
'                           "Amount,Rol,Posted,Remarks,CPosted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbBudgetHead.Columns(0).text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & "," & IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & "," & IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 12)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
'
''        End If
'
'
''  --------------------------------Posting Information-----------------------------------------
'
'
'
'   ElseIf cmdPost.Caption = "&Posting" Then
'
'
''     Dim iPost
'     txtpost.text = "Posted"
'
'
'
'ipost = MsgBox("Do you want to Post this bill?", vbYesNo)
'
'If ipost = vbYes Then
'
'     txtpost.text = "Posted"
'     cn.Execute "UPDATE SReqMaster SET  ReqDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              " " & _
'                              "DeptName='" & cmbBudgetHead.Columns(0).text & "',CPost='" & txtpost & "', UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SReqDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'        j = 0
'            For j = 1 To fgStock.Rows - 1
'
'            If fgStock.Cell(flexcpChecked, j, 11) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'''                If fgExport.TextMatrix(j, 1) <> "" Then
'                cn.Execute "INSERT INTO SReqDetails (SerialNo,ReqDate,DeptName,Catagory,SubCatagory,ItemName,Qty,Rate, " & _
'                           "Amount,Rol,Posted,Remarks,CPosted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbBudgetHead.Columns(0).text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & "," & IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & "," & IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 12)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
'
'        End If
'    End If
'Else
'   cn.BeginTrans
''     flagSlNo = 0
'    If CmdNew.Caption = "&Save" Then
'
'    'General Information for Payment Master
'     strSQL = "INSERT INTO SReqMaster (ReqDate,DeptName, CPost,UserName " & _
'                ") " & _
'                "VALUES ('" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                " '" & cmbBudgetHead.Columns(0).text & "','" & txtpost & "','" & txtUserName.text & "')"
'     cn.Execute strSQL
'      rcupdate = True
''     cn.CommitTrans
'
''     -------------For primary key and foreign key relation------------
'         If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(SerialNo),1) as InvNo from SReqMaster"
'           rs.Open str, cn, adOpenStatic, adLockReadOnly
'           txtSerialNo = Val(rs!InvNo)
'
''------------------------
'
''     '" & parseQuotes(txtSerialNo) & "',
'    'payment Detail Information Enter This table
''    strDeliveryDate = "'" & parseQuotes(fgStock.TextMatrix(j, 11)) & "'"
'            j = 0
'            For j = 1 To fgStock.Rows - 1
'
'            If fgStock.Cell(flexcpChecked, j, 11) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'''                If fgExport.TextMatrix(j, 1) <> "" Then
'                cn.Execute "INSERT INTO SReqDetails (SerialNo,ReqDate,DeptName,Catagory,SubCatagory,ItemName,Qty,Rate, " & _
'                           "Amount,Rol,Posted,Remarks,CPosted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbBudgetHead.Columns(0).text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & "," & IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & "," & IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 12)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "')"
'               Next
'        rcupdate = True
'
'
'
'        cn.CommitTrans
'        MsgBox "Record added Successfully", vbInformation, "Confirmation"
'
'    ' Update Information
'
'
'
'ElseIf (CmdEdit.Caption = "&Update") Then
'
''            If txtpost.text = "Not Posted" Then
'
'                   cn.Execute "UPDATE SReqMaster SET  ReqDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              " " & _
'                              "DeptName='" & cmbBudgetHead.Columns(0).text & "',CPost='" & txtpost & "', UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SReqDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'         j = 0
'            For j = 1 To fgStock.Rows - 1
'
'            If fgStock.Cell(flexcpChecked, j, 11) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'''                If fgExport.TextMatrix(j, 1) <> "" Then
'                cn.Execute "INSERT INTO SReqDetails (SerialNo,ReqDate,DeptName,Catagory,SubCatagory,ItemName,Qty,Rate, " & _
'                           "Amount,Rol,Posted,Remarks,CPosted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbBudgetHead.Columns(0).text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & "," & IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & "," & IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 12)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
'
''        End If
'
'
''  --------------------------------Posting Information-----------------------------------------
'
'
'
'   ElseIf cmdPost.Caption = "&Posting" Then
'
'
''     Dim iPost
'     txtpost.text = "Posted"
'
'
'
'ipost = MsgBox("Do you want to Post this bill?", vbYesNo)
'
'If ipost = vbYes Then
'
'     txtpost.text = "Posted"
'     cn.Execute "UPDATE SReqMaster SET  ReqDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              " " & _
'                              "DeptName='" & cmbBudgetHead.Columns(0).text & "',CPost='" & txtpost & "', UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SReqDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'        j = 0
'            For j = 1 To fgStock.Rows - 1
'
'            If fgStock.Cell(flexcpChecked, j, 11) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'''                If fgExport.TextMatrix(j, 1) <> "" Then
'                cn.Execute "INSERT INTO SReqDetails (SerialNo,ReqDate,DeptName,Catagory,SubCatagory,ItemName,Qty,Rate, " & _
'                           "Amount,Rol,Posted,Remarks,CPosted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbBudgetHead.Columns(0).text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & "," & IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & "," & IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 12)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
'
'        End If
'' End If
''----------------------Undo Posting------
'  ElseIf cmdUndoPost.Caption = "&Undo Posting" Then
'
'
''     Dim iPost
'    txtpost.text = "Not Posted"
'
'
'
'ipost = MsgBox("Do you want to Undo Post this bill?", vbYesNo)
'
'If ipost = vbYes Then
'
'     txtpost.text = "Not Posted"
'     cn.Execute "UPDATE SReqMaster SET  ReqDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              " " & _
'                              "DeptName='" & cmbBudgetHead.Columns(0).text & "',CPost='" & txtpost & "', UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SReqDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'        j = 0
'            For j = 1 To fgStock.Rows - 1
'
'            If fgStock.Cell(flexcpChecked, j, 11) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'''                If fgExport.TextMatrix(j, 1) <> "" Then
'                cn.Execute "INSERT INTO SReqDetails (SerialNo,ReqDate,DeptName,Catagory,SubCatagory,ItemName,Qty,Rate, " & _
'                           "Amount,Rol,Posted,Remarks,CPosted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbBudgetHead.Columns(0).text & "','" & parseQuotes(fgStock.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgStock.TextMatrix(j, 5)) & "','" & parseQuotes(fgStock.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgStock.TextMatrix(j, 7) = "", "0", fgStock.TextMatrix(j, 7)) & "," & IIf(fgStock.TextMatrix(j, 8) = "", "0", fgStock.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgStock.TextMatrix(j, 9) = "", "0", fgStock.TextMatrix(j, 9)) & "," & IIf(fgStock.TextMatrix(j, 10) = "", "0", fgStock.TextMatrix(j, 10)) & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgStock.TextMatrix(j, 12)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgStock.TextMatrix(j, 14)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record undo Posted Successfully", vbInformation, "Confirmation"
'
'        End If
' End If
'
''-----------------------End Undo Posting---
'
'End If
''End If
''    cn.CommitTrans
'
'    Exit Function
'
'''''''ErrHandler:
'''''''
'''''''    cn.RollbackTrans
'''''''    Select Case Err.Number
'''''''        Case -2147217900
'''''''            MsgBox "Please select Numeric number in ROL field", vbInformation, "Confirmation"
'''''''
'''''''   End Select
'
'
''   If Err.Number = -2147217874 Then
''    MsgBox "You can't Insert same item from same style multiple times in one BTB LC."
'''   End If
''            MsgBox cn.Errors(0).NativeError & " : " & cn.Errors(0).Description
''    End Select
'End Function
'
'Private Function IsValidRecord() As Boolean
'    IsValidRecord = True
'
'    If Trim(ReceivingDate) = "" Then
'        MsgBox "Your are missing Receiving Information.", vbInformation
'        ReceivingDate.SetFocus
'        IsValidRecord = False
'        Exit Function
'
'
'  ElseIf Trim(cmbBudgetHead) = "" Then
'        MsgBox "Your are missing Department Name Information.", vbInformation
''        cmbSupplierName.SetFocus
'        IsValidRecord = False
'        Exit Function
'
''    ---------------------------------------------------
'
''ElseIf cmdNew.Caption = "&Save" Or cmdEdit.Caption = "&Update" Then
''         Dim k As Integer
''         If rsItemDetail.RecordCount > 0 Then
''         If rsItemDetail.State <> 0 Then rsItemDetail.Close
''            rsItemDetail.Open "select * from tblItemDetail where ItemCode='" & fgItem.TextMatrix(Row, 4) & "'", cn, adOpenStatic, adLockReadOnly
''
''             If Not rsItemDetail.EOF Then
''        MsgBox "This Record exists Duplicate ItemCode No.", vbInformation, Me.Caption & " - " & App.Title
''          fgItem.TextMatrix(k, 4).SetFocus
''          IsValidRecord = False
''         Exit Function
''            End If
''         End If
''         End If
''    Exit Function
''-----------------------------------------------------------------------
'
''-----------------------------------------------------------------------
'    Else
'
''        Dim j As Integer
''
''         For j = 1 To fgItem.Rows - 2
''
''        If Not IsNumeric(fgItem.TextMatrix(j, 6)) Then
''        MsgBox "Select Numeric value in ROL field.", vbInformation
'''         fgItem.TextMatrix(j, 4) = ""
'''         fgItem.RemoveItem fgItem.Row
''        IsValidRecord = False
''
''        End If
''
''       Next
'
'       Exit Function
'     End If
'    End Function
'
'Private Sub FindRecord()
'
'    Dim i As Integer
'    Dim strPaymentDetail As String
'    Set rsItemDetail = New ADODB.Recordset
'    txtSerialNo = rsItemMaster!SerialNo
''    cmbSupplierName = rsItemMaster!SupplierName
'    ReceivingDate = rsItemMaster!ReqDate
''    txtSupplierBill = rsItemMaster!SupplierBill
'    cmbBudgetHead = rsItemMaster!DeptName
''    chkAutoposting = rsItemMaster!Autoposting
'    txtpost = rsItemMaster!Cpost
'    txtUserName = rsItemMaster!UserName
'
'
'    fgStock.Rows = 1
'    strPaymentDetail = "SELECT  SerialNo, ReqDate, DeptName,Catagory ,SubCatagory,ItemName,Qty, " & _
'                "Rate,Amount,Rol,Posted,Remarks,CPosted,Unit FROM SReqDetails " & _
'                "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "
'    rsItemDetail.CursorLocation = adUseClient
'    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly
'
'
'    If rsItemDetail.RecordCount <> 0 Then
'
'        fgStock.Rows = rsItemDetail.RecordCount + 1
''                i = 0
'        For i = 1 To rsItemDetail.RecordCount
'            fgStock.TextMatrix(i, 1) = rsItemDetail("SerialNo")
'            fgStock.TextMatrix(i, 2) = rsItemDetail("ReqDate")
'            fgStock.TextMatrix(i, 3) = rsItemDetail("DeptName")
'            fgStock.TextMatrix(i, 4) = rsItemDetail("Catagory")
'            fgStock.TextMatrix(i, 5) = rsItemDetail("SubCatagory")
'            fgStock.TextMatrix(i, 6) = rsItemDetail("ItemName")
'            fgStock.TextMatrix(i, 7) = rsItemDetail("Qty")
'            fgStock.TextMatrix(i, 8) = rsItemDetail("Rate")
'            fgStock.TextMatrix(i, 9) = rsItemDetail("Amount")
'            fgStock.TextMatrix(i, 10) = rsItemDetail("Rol")
'            fgStock.TextMatrix(i, 11) = rsItemDetail("Posted")
'            fgStock.TextMatrix(i, 12) = rsItemDetail("Remarks")
'            fgStock.TextMatrix(i, 13) = rsItemDetail("CPosted")
'            fgStock.TextMatrix(i, 14) = rsItemDetail("Unit")
'        rsItemDetail.MoveNext
'        Next
'      End If
'        rsItemDetail.Close
'End Sub
'
'
'Public Sub printReport()
'
'On Error GoTo ErrH
'    Dim strPath    As String
'    Dim strSQL     As String
'    Dim temp       As Double
'    If rsItemMaster.RecordCount = 0 Then
'        MsgBox "Data not available", vbInformation, "Confarmation"
'        Exit Sub
'    End If
'
'
'        strPath = App.Path + "\reports\ReqReceipt.rpt"
'        Set objReportApp = CreateObject("Crystal.CRPE.Application")
'        Set objReport = objReportApp.OpenReport(strPath)
'        Set objReportDatabase = objReport.Database
'        Set objReportDatabaseTables = objReportDatabase.Tables
'        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
'        Set ObjPrinterSetting = objReport.PrintWindowOptions
'        Set objReportFormulaFieldDefinations = objReport.FormulaFields
'
'
'
'    Set rsDailyRpt = New ADODB.Recordset
'If rsDailyRpt.State <> 0 Then rsDailyRpt.Close
'
'
'
'
'            strSQL = "SELECT SReqMaster.SerialNo, SReqMaster.ReqDate, " & _
'                      "SReqMaster.DeptName,SReqMaster.UserName,SReqDetails.Catagory, SReqDetails.SubCatagory, SReqDetails.ItemName, " & _
'                      "SReqDetails.Qty,SReqDetails.Rate, SReqDetails.Amount, " & _
'                      "SReqDetails.Posted ,SReqDetails.Remarks,SReqDetails.Unit " & _
'                      "FROM SReqMaster INNER JOIN " & _
'                      "SReqDetails ON SReqMaster.SerialNo = SReqDetails.SerialNo and SReqMaster.SerialNo ='" & Me.txtSerialNo & "' ORDER BY SReqDetails.Catagory "
'
'                      rsDailyRpt.Open strSQL, cn, adOpenStatic
'
''strSQL = "SELECT SReqMaster.SerialNo, SReqMaster.ReqDate, " & _
''                      "SReqMaster.DeptName,SReqMaster.UserName,SReqDetails.Catagory, SReqDetails.SubCatagory, SReqDetails.ItemName, " & _
''                      "SReqDetails.Qty,SReqDetails.Rate, SReqDetails.Amount, " & _
''                      "SReqDetails.Posted ,SReqDetails.Remarks,SReqDetails.Unit " & _
''                      "FROM SReqMaster,SReqDetails WHERE " & _
''                      "SReqMaster.SerialNo ='" & Me.txtSerialNo & "' ORDER BY SReqDetails.SerialNo "
'
'
''                      rsDailyRpt.Open strSQL, cn, adOpenStatic
''        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
''            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"
'
'
'        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt
'
'        ObjPrinterSetting.HasPrintSetupButton = True
'        ObjPrinterSetting.HasRefreshButton = True
'        ObjPrinterSetting.HasSearchButton = True
'        ObjPrinterSetting.HasZoomControl = True
'
'        objReport.DiscardSavedData
'        objReport.Preview "Menu Item List Report", , , , , 16777216 Or 524288 Or 65536
'
'
'        Set objReport = Nothing
'        Set objReportDatabase = Nothing
'        Set objReportDatabaseTables = Nothing
'        Set objReportDatabaseTable = Nothing
'    Exit Sub
'
'ErrH:
'
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "You need to select record first", vbInformation, "Item Catagory List Report"
'        Case Else
'            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Item catagory Report"
'    End Select
'End Sub
'
'
'Private Sub duplicate()
'   Dim j As Integer
'
'         For j = 1 To fgStock.Rows - 2
'
'        If Val(fgStock.TextMatrix(j, 4)) = Val(fgStock.TextMatrix(j + 1, 4)) Then
'        MsgBox "Duplicate Item Code Number.", vbInformation
'         fgStock.TextMatrix(j, 4) = ""
'         End If
'
'         Next
'
'End Sub
'
'Public Sub PopulateForm(StrID As String)
'    rsItemMaster.Close
'    rsItemMaster.Open "select * from SReqMaster", cn, adOpenStatic, adLockReadOnly
'    rsItemMaster.MoveFirst
'    rsItemMaster.Find "SerialNo=" & parseQuotes(StrID)
'    If rsItemMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord
'
'End Sub
'
'
'Private Sub fgStock_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'
''         Dim k As Integer
''
''         Set rsItemDetail = New ADODB.Recordset
''
''                 If rsItemDetail.State <> 0 Then rsItemDetail.Close
''         If col = 5 Then
''            rsItemDetail.Open "select * from SMSSubCatagoryDetail where SubCatagoryCode='" & fgStock.TextMatrix(row, 4) & "'", cn, adOpenStatic, adLockReadOnly
''
''             If Not rsItemDetail.EOF Then
''        MsgBox "This Record exists Duplicate Item Code No.", vbInformation, Me.Caption & " - " & App.Title
''            End If
''         End If
''
''
''' .-------------------------------
''
''    If col = 6 Then
''          Dim j As Integer
''
''        For j = 1 To fgStock.Rows - 1
''
''        If Not IsNumeric(fgStock.TextMatrix(j, 6)) Then
''        MsgBox "Select Numeric value in ROL field.", vbInformation
''
''        End If
''
''       Next
'' End If
''
''
''' ---------------------------------------
''
''
''If fgStock.Rows > 2 Then
''        For j = 1 To fgStock.Rows - 1
''            If fgStock.TextMatrix(j, 5) = fgStock.TextMatrix(row, 5) And j <> fgStock.row Then
''                MsgBox "This charge already selected.", vbInformation
''                fgStock.TextMatrix(row, 5) = ""
''            End If
''        Next
''    End If
''
'''               --------------------------
''
''
''
''End Sub
''
'''Private Sub fgStock_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'''Dim j As Integer
'''
'''If fgStock.Rows > 2 Then
'''        For j = 1 To fgStock.Rows - 1
'''            If fgStock.TextMatrix(j, 5) = fgStock.TextMatrix(Row, 5) And j <> fgStock.Row Then
'''                MsgBox "This charge already selected.", vbInformation
'''                fgStock.TextMatrix(Row, 5) = ""
'''            End If
'''        Next
'''    End If
'''End Sub
''
''
''
''Private Sub check()
''
''Dim i As Integer
''Dim j As Integer
''Dim sSQL As String
''For i = fgStock.Rows - 1 To 1 Step -1
''  sSQL = ""
''  For j = 0 To fgStock.Cols - 1
''     sSQL = sSQL & Trim(fgStock.TextMatrix(i, j))
''  Next
''  If Trim(sSQL) = "" Then
''      fgStock.RemoveItem i
''  End If
''  If fgStock.Rows <= 1 Then Exit For
''Next
''
''
''
'''    If Col = 5 Then
'''        If CDbl(fgStock.TextMatrix(Row, 11)) > CDbl(fgStock.TextMatrix(Row, 10)) Then
'''            MsgBox "Rejected quantity can't be greater than Receive Quantity.", vbInformation
'''            fgStock.Select Row, 11
'''        End If
'''    End If
''End Sub
''
''
''
'''Private Sub deleteRow()
'''If fgStock.Rows = 1 Then fgStock.AddItem ""
'''        On Error Resume Next
'''        If fgStock.Rows <= 1 Then Exit Sub
'''Dim i As Integer
'''Dim j As Integer
'''    For i = 0 To fgStock.Rows - 1
'''             j = 1
'''             For j = 1 To fgStock.Rows - 1
'''
'''                If fgStock.TextMatrix(i, 5) = fgStock.TextMatrix(j, 5) Then
'''                    fgStock.RemoveItem j
'''                End If
'''             Next
'''    Next
'''
'''End Sub
''
''
'
''----------------------------------------------------------------------------------
'
''Option Explicit
''
''Private Sub Check1_Click()
'''Dim iRows As Integer
'''Dim i As Integer
''' For i = 1 To fgStock.FixedRows - 1
'''       fgStock.Cell(flexcpChecked, i, 10) = flexChecked
'''       Next
''      Dim f As Integer
''      f = fgStock.Rows - 1
''  If Check1.Value = 1 Then
''
'''f = fgStock.Rows - 1
''Dim i As Integer
''For i = 1 To f
''    fgStock.Cell(flexcpChecked, i, 10) = flexChecked
'''    fgStock.TextMatrix(i, 2) = "Descripción"
'''    fgStock.TextMatrix(i, 3) = "Precio"
''Next i '
'''Dim f As Integer
'''f = fgStock.Rows - 1
'''Dim i As Integer
''Else
''For i = 1 To f
''    fgStock.Cell(flexcpChecked, i, 1, i, 1) = flexUnchecked
'''    fgStock.TextMatrix(i, 2) = "Descripción"
'''    fgStock.TextMatrix(i, 3) = "Precio"
''    Next i
''
''
''End If
''
''
''
''End Sub
''
''Private Sub cmdAddItem_Click()
''frmItemInputReceiving.Show vbModal
''End Sub
''
''Private Sub cmdNew_Click()
''' Dim r&
'''
'''        For r = fgStock.FixedRows To fgStock.Rows - 1
'''
'''            If fgStock.Cell(flexcpChecked, r, 1) = flexChecked Then
'''
'''                Debug.Print fgStock.TextMatrix(r, 1); " is Checked"
'''
'''            End If
'''
'''        Next
''
''Dim f As Integer
''f = fgStock.Rows - 1
''Dim i As Integer
''For i = 1 To f
''    fgStock.Cell(flexcpChecked, i, 1, i, 1) = flexChecked
'''    fgStock.TextMatrix(i, 2) = "Descripción"
'''    fgStock.TextMatrix(i, 3) = "Precio"
''Next i
''
''
''End Sub
''
''Private Sub fgStock_Click()
''fgStock.Editable = flexEDKbdMouse
''End Sub
''
''Private Sub Form_Load()
'' fgStock.ColDataType(10) = flexDTBoolean
''' fgStock.ColDataType(1) = flexDTBoolean
'End Sub
'
'
