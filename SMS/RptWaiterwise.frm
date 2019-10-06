VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form rptwaiterbysales 
   BackColor       =   &H00C0B4A9&
   Caption         =   " Waiter wise Sales Report"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6675
   Icon            =   "RptWaiterwise.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDateSelect 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Date Range"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5535
      Begin VB.OptionButton OpCurrentDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&rrent Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton OpCustomDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&stom Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1680
         Top             =   3840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptWaiterwise.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptWaiterwise.frx":0E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptWaiterwise.frx":173E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   2280
         TabIndex        =   6
         Top             =   3840
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   1058
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Preview"
               Object.ToolTipText     =   "Preview"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Close"
               Object.ToolTipText     =   "Close"
               ImageIndex      =   3
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbWaiter 
         Height          =   345
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   2520
         AllowInput      =   0   'False
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   4445
         _ExtentY        =   609
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Label lblWaiter 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Waiter Name"
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
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H00C0B4A9&
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   2040
         TabIndex        =   7
         Top             =   2880
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc dcSWaiter 
      Height          =   450
      Left            =   3720
      Top             =   4080
      Visible         =   0   'False
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   794
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
      Caption         =   "Waiter"
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
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Caption         =   "Waiter Wise Sales Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "rptwaiterbysales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
'
'Private rsMaster                            As ADODB.Recordset
'Private rsSelect                            As ADODB.Recordset 'sub
'
'Private objReportApp                        As CRPEAuto.Application
'Private objReport                           As CRPEAuto.Report
'Private objReportDatabase                   As CRPEAuto.Database
'Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
'Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
'Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private objReportFF                         As CRPEAuto.FormulaFieldDefinition
'Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
'Private Tracer                              As Integer
'Private strGroupName                        As String
'
'Private Sub CmdExit_Click()
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Call Connect
'    Call WaiterName
'    ModFunction.StartUpPosition Me
'    OpCurrentDate.Visible = True
'    OpCustomDate.Visible = True
''    SSCurrentDate.Date = Date
''    SSTDate.Date = Date
'
'End Sub
'
'Private Sub WaiterName()
'     dcSWaiter.CursorLocation = adUseClient
'     dcSWaiter.ConnectionString = cn.ConnectionString
'     dcSWaiter.LockType = adLockReadOnly
'     dcSWaiter.RecordSource = "SELECT  WaiterName,WaiterRemaks FROM tblWaiterName ORDER BY WaiterName"
'     cmbWaiter.DataMode = ssDataModeBound
'     Set cmbWaiter.DataSource = dcSWaiter
'     cmbWaiter.DataSourceList = dcSWaiter
'     cmbWaiter.DataFieldList = "WaiterName"
'     cmbWaiter.DataField = "WaiterName"
'     cmbWaiter.ColumnHeaders = True
'     cmbWaiter.BackColorOdd = &HFFFF00
'     cmbWaiter.BackColorEven = &HFFC0C0
'     cmbWaiter.ForeColorEven = &H80000008
'
'End Sub
'
'
'Private Sub OpCurrentDate_Click()
'    OpCurrentDate.Visible = True
'    OpCustomDate.Visible = True
''    SSCurrentDate.Visible = True
'    lblFrom.Visible = False
''    SSFDate.Visible = False
'    lblTo.Visible = False
''    SSTDate.Visible = False
'End Sub
'Private Sub opCustomDate_Click()
'    OpCustomDate.Visible = True
'    OpCurrentDate.Visible = True
'    lblFrom.Visible = True
'    lblTo.Visible = True
''    SSCurrentDate.Visible = False
''    SSFDate.Visible = True
''    SSTDate.Visible = True
'End Sub
'Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
'  Select Case Button.Key
'     Case "Preview"
'            If Validate Then
'                Tracer = 0
'                Call FetchData
'                Call previewReport
'               End If
'     Case "Print"
'            If Validate Then
'                Tracer = 1
'                Call FetchData
'                Call previewReport
'               End If
'     Case "Close"
'               Unload Me
'    End Select
'
'End Sub
'Private Function Validate() As Boolean
'           Validate = True
'        If SSFDate.Date > SSTDate.Date Then
'            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
'            SSFDate.SetFocus
'            Validate = False
'            Exit Function
'        End If
'    End Function
'
'Public Function parseQuotes(text As String) As String
'    parseQuotes = Replace(text, "'", "''")
'End Function
'
'Public Function FetchData()
'
'    Set rsMaster = New ADODB.Recordset
'
'    If OpCurrentDate.Value = True Then
'
'
'               rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate, " & _
'                         "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
'                         "tblCashDetail.ItemCode,tblCashDetail.ItemName,tblCashDetail.Qty,tblCashDetail.Rate, " & _
'                         "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount " & _
'                         "FROM tblCashMaster INNER JOIN " & _
'                         "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate='" & SSCurrentDate.Date & "' " & _
'                         "And tblCashMaster.CActive = 'Active' And tblCashMaster.CPost = 'Posted' and tblCashMaster.WaiterName='" & parseQuotes(cmbWaiter) & "'", cn, adOpenStatic, adLockReadOnly
'
'
'
''            rsMaster.Open "SELECT USSalesMaster.InvoiceNo, USSalesDetail.Name, " & _
''                         "USSalesDetail.IssueQty, USSalesDetail.UnitPrice, " & _
''                         "USSalesDetail.Total, USSalesMaster.Date, " & _
''                         "USSalesMaster.SoldBy , USSalesDetail.Remarks " & _
''                         "FROM USSalesMaster INNER JOIN " & _
''                         "USSalesDetail ON " & _
''                         "USSalesMaster.InvoiceNo = USSalesDetail.InvoiceNo where " & _
''                         "USSalesMaster.Date='" & DTPicker1.Value & "'", cn, adOpenStatic, adLockReadOnly
'
'     End If
'
'      If OpCustomDate.Value = True Then
'
'       rsMaster.Open "SELECT tblCashMaster.SerialNo,tblCashMaster.strDate, " & _
'                         "ItemGroup=(SELECT ItemGroup from tblItemMaster where tblItemMaster.SerialNo=tblCashDetail.SerialNo), " & _
'                         "tblCashDetail.ItemCode,tblCashDetail.ItemName,tblCashDetail.Qty,tblCashDetail.Rate, " & _
'                         "(tblCashDetail.Qty*tblCashDetail.Rate) as Amount " & _
'                         "FROM tblCashMaster INNER JOIN " & _
'                         "tblCashDetail ON tblCashMaster.SerialNo=tblCashDetail.BillSerialNo AND tblCashMaster.strDate  " & _
'                         "BETWEEN '" & SSFDate.Date & "' AND '" & SSTDate.Date & "' And tblCashMaster.CActive = 'Active' " & _
'                         "And tblCashMaster.CPost = 'Posted' and tblCashMaster.WaiterName='" & parseQuotes(cmbWaiter) & "'", cn, adOpenStatic, adLockReadOnly
'
'      End If
'
'End Function
'
'
'Public Sub previewReport()
'On Error GoTo ErrH
'    Dim strPath As String
'
'    If rsMaster.RecordCount = 0 Then
'        MsgBox "Data not available", vbInformation
'        Exit Sub
'    End If
'
'
'        strPath = App.Path + "\reports\WaiterWiseSales.rpt"
'        Set objReportApp = CreateObject("Crystal.CRPE.Application")
'        Set objReport = objReportApp.OpenReport(strPath)
'        Set objReportDatabase = objReport.Database
'        Set objReportDatabaseTables = objReportDatabase.Tables
'        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
'        Set ObjPrinterSetting = objReport.PrintWindowOptions
'        Set objReportFormulaFieldDefinations = objReport.FormulaFields
'
'   If OpCurrentDate.Value = True Then
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'              objReportFF.text = "'" + Format(SSCurrentDate, "dd-MMM-yyyy") + "'"
'   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
'              objReportFF.text = "'" + cmbWaiter + "'"
'
'
'  End If
'
'
'  If OpCustomDate.Value = True Then
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
'
'              objReportFF.text = "'" + Format(SSFDate, "dd-MMM-yyyy") + "'"
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
'             objReportFF.text = "'" + Format(SSTDate, "dd-MMM-yyyy") + "'"
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
'              objReportFF.text = "'" + cmbWaiter + "'"
'
'
'   End If
'
'
'
'        objReportDatabaseTable.SetPrivateData 3, rsMaster
'
'        ObjPrinterSetting.HasPrintSetupButton = True
'        ObjPrinterSetting.HasRefreshButton = True
'        ObjPrinterSetting.HasSearchButton = True
'        ObjPrinterSetting.HasZoomControl = True
'
'        objReport.DiscardSavedData
'        objReport.Preview "Waiter Insformations", , , , , 16777216 Or 524288 Or 65536
'
'
'     If Tracer = 1 Then
'    objReport.PrintOut
'    End If
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
'        Case 20545
'            MsgBox "Request cancelled by the user", vbInformation, "Waiter Information Report"
'        Case Else
'            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Waiter Information Report"
'    End Select
'End Sub
'
'

