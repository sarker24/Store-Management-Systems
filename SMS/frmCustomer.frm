VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCustomer 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Informations"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   375
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Find Last"
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Find Next"
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Find Previous"
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Find First"
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   6120
      Picture         =   "frmCustomer.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   795
      Left            =   4200
      Picture         =   "frmCustomer.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      Width           =   945
   End
   Begin VB.TextBox txtCName 
      Height          =   495
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   4815
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   7095
      Begin VB.TextBox txtCAddress 
         Height          =   1755
         Left            =   1680
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1440
         Width           =   5175
      End
      Begin VB.TextBox txtCPhone 
         Height          =   465
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   2
         Text            =   " "
         Top             =   3240
         Width           =   5175
      End
      Begin VB.TextBox txtCFax 
         Height          =   465
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   3720
         Width           =   5175
      End
      Begin VB.TextBox txtCEmail 
         Height          =   465
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Text            =   " "
         Top             =   4215
         Width           =   5175
      End
      Begin VB.TextBox txtCID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   465
         Left            =   1680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Top             =   240
         Width           =   2655
      End
      Begin VB.CheckBox chkActive 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Left            =   5760
         Top             =   240
      End
      Begin VB.Label lblCustomerName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier Name"
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
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblCustomerAddress 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier Address"
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
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblCustomerID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblCustomerPhone 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblCustomerAddress 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier Fax"
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
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label lblCustomerAddress 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier E-mail"
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
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   4215
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      Height          =   795
      Left            =   2280
      Picture         =   "frmCustomer.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6000
      Width           =   945
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   360
      Picture         =   "frmCustomer.frx":1FE8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6000
      Width           =   945
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1320
      Picture         =   "frmCustomer.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   3240
      Picture         =   "frmCustomer.frx":317C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   945
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      Height          =   795
      Left            =   5160
      Picture         =   "frmCustomer.frx":3A46
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      Caption         =   " Supplier Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   7335
   End
   Begin MSAdodcLib.Adodc DCRSearch 
      Height          =   330
      Left            =   1560
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private rsfactory             As ADODB.Recordset
'Private strFileName           As String
'Private bRecordExists         As Boolean
'Private rm                    As New ADODB.Recordset
'Private rs                    As New ADODB.Recordset
'Dim str As String
''--------------------------------------------------------------
'Private oReportApp                        As CRPEAuto.Application
'Private oReport                           As CRPEAuto.Report
'Private oReportDatabase                   As CRPEAuto.Database
'Private oReportDatabaseTables             As CRPEAuto.DatabaseTables
'Private oReportDatabaseTable              As CRPEAuto.DatabaseTable
''Private oReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
''Private oReportFF                         As CRPEAuto.FormulaFieldDefinition
'Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions
'
'Private Sub cmdDelete_Click()
'On Error GoTo ErrHandler
'     Dim idelete As Integer
'     idelete = MsgBox("Do you want to delete this record?", vbYesNo)
'     If frmLogin.txtUName.text = "Admin" Then
'    If idelete = vbYes Then
'
'    cn.Execute "Delete From SSuplierName Where sID ='" & parseQuotes(txtCID) & "'"
'            Call allClear
'
''           move Next
'    End If
'
'    End If
'ErrHandler:
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
'     End Select
'End Sub
'
'Private Sub CmdPreview_Click()
'    Call printReport
'End Sub
'
''Private Sub cmdPreview_Click()
''    Call printReport
''End Sub
'
'Private Sub cmdClose_Click()
'    Unload Me
'End Sub
'Private Sub cmdCancel_Click()
'    cmdCancel.Enabled = False
'    cmdNew.Enabled = True
'    cmdNew.Caption = "&New"
'    cmdEdit.Caption = "&Edit"
'    cmdPreview.Enabled = True
'    CmdDelete.Enabled = True
'    cmdOpen.Enabled = True
'    cmdClose.Enabled = True
'    cmdEdit.Enabled = True
'    txtCID.Enabled = False
'    Call allClear
'    Call alldisable
'    If Not rsfactory.EOF Then FindRecord
'End Sub
'
''
'Private Sub cmdNew_Click()
'    On Error GoTo ProcError
'      Set rs = New ADODB.Recordset
'    If cmdNew.Caption = "&New" Then
'        cmdNew.Caption = "&Save"
'        cmdEdit.Enabled = False
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        CmdDelete.Enabled = False
'        cmdOpen.Enabled = False
'        cmdPreview.Enabled = False
''        chkActive.Enabled = False
'        Call allClear
'
'If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(sID),0) as SerialNo from SSuplierName"
'           rs.Open str, cn, adOpenStatic, adLockReadOnly
'           txtCID.text = Val(rs!SerialNo) + 1
'
'        Call allenable
'        txtCName.SetFocus
'    ElseIf cmdNew.Caption = "&Save" Then
'        Dim s As String
'        If IsValidRecord Then
'            If rcupdate Then
'                txtCID.Enabled = False
'                cmdNew.Caption = "&New"
'                cmdEdit.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                CmdDelete.Enabled = True
'                cmdOpen.Enabled = True
'                cmdPreview.Enabled = True
''                chkActive.Enabled = True
'                Call alldisable
'                s = txtCName
'                rsfactory.Requery
'                rsfactory.MoveFirst
'                rsfactory.Find "CName='" & parseQuotes(s) & "'"
'                FindRecord
'
'            End If
'        End If
'    End If
''
'    Exit Sub
'
'ProcError:
'    Select Case Err.Number
'    Case 0:
'    Case Else
'        MsgBox Err.Description
'    End Select
'
'End Sub
'
'Private Sub cmdEdit_Click()
'    If cmdEdit.Caption = "&Edit" Then
'        cmdNew.Enabled = False
'        Call allenable
'        txtCName.SetFocus
'        cmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        CmdDelete.Enabled = False
'        cmdPreview.Enabled = False
'        cmdOpen.Enabled = False
''        chkActive.Enabled = False
'ElseIf cmdEdit.Caption = "&Update" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdEdit.Caption = "&Edit"
'                cmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                CmdDelete.Enabled = True
''                chkActive.Enabled = True
'        cmdPreview.Enabled = True
'        cmdOpen.Enabled = True
'                Call alldisable
'                rsfactory.Requery
'
'                Dim s As String
'                s = txtCName
'                rsfactory.Find "CName='" & parseQuotes(s) & "'"
''                Call search
''                Call countrysearch
'                FindRecord
'
'            End If
'        End If
'    End If
'End Sub
'
'Private Sub cmdOpen_Click()
'    frmCustomerSearch.Show vbModal
'    cmdOpen.Enabled = True
'    cmdCancel.Enabled = True
'End Sub
'
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
'End Sub
'
'Private Sub Form_Load()
'
'    Call Connect
'       ModFunction.StartUpPosition Me
'    Set rsfactory = New ADODB.Recordset
'    rsfactory.Open "select * from SSuplierName", cn, adOpenStatic, adLockReadOnly
'    Call alldisable
'   If rsfactory.RecordCount > 0 Then
'        bRecordExists = True
'    Else
'        bRecordExists = False
'    End If
'
'    If Not rsfactory.EOF Then FindRecord
'
'    txtCID.Enabled = False
'
'End Sub
'
'Private Sub allenable()
''    txtDiscountCard.Enabled = True
'    txtCName.Enabled = True
'    txtCAddress.Enabled = True
'    txtCPhone.Enabled = True
'    txtCFax.Enabled = True
'    txtCEmail.Enabled = True
'    chkActive.Enabled = True
'End Sub
'
'Private Sub alldisable()
''    txtDiscountCard.Enabled = False
'    txtCName.Enabled = False
'    txtCAddress.Enabled = False
'    txtCPhone.Enabled = False
'    txtCFax.Enabled = False
'    txtCEmail.Enabled = False
'    chkActive.Enabled = False
'End Sub
'
'Private Sub allClear()
''    txtDiscountCard.text = ""
'    txtCName.text = ""
'    txtCAddress.text = ""
'    txtCPhone.text = ""
'    txtCFax.text = ""
'    txtCEmail.text = ""
'    chkActive.Value = 0
'End Sub
'
'Private Function rcupdate() As Boolean
''    On Error GoTo ErrHandler
'    cn.BeginTrans
'    If cmdNew.Caption = "&Save" Then
'
'    cn.Execute "INSERT INTO SSuplierName(sID,sName,sAddress,sPhone, " & _
'                   " sProvince,sEmail,Active) " & _
'                   " VALUES ('" & parseQuotes(txtCID) & "','" & parseQuotes(txtCName) & "','" & parseQuotes(txtCAddress) & "', " & _
'                   " '" & parseQuotes(txtCPhone) & "', " & _
'                   " '" & parseQuotes(txtCFax) & "', " & _
'                   " '" & parseQuotes(txtCEmail) & "', " & _
'                   " '" & parseQuotes(chkActive) & "') "
'
'
'          rcupdate = True
'          MsgBox "Record Added", vbInformation, "Confirmation"
'    Else
'
'
'cn.Execute "Update SSuplierName Set sName='" & parseQuotes(txtCName) & _
'                  "',sAddress='" & parseQuotes(txtCAddress) & "',sPhone='" & parseQuotes(txtCAddress) & "', " & _
'                  " CPhone='" & parseQuotes(txtCPhone) & "',CFax='" & parseQuotes(txtCFax) & "', " & _
'                  "',CEmail='" & parseQuotes(txtCEmail) & "',Active='" & parseQuotes(chkActive) & "' " & _
'                  " Where sID ='" & parseQuotes(txtCID) & "' "
'
'        rcupdate = True
'        MsgBox "Record Updated", vbInformation, "Confirmation"
'    End If
'
'    cn.CommitTrans
''    Exit Sub
'    Exit Function
'
'
'
''ErrHandler:
''    cn.RollbackTrans
''   ' rsFactory.Requery
''    Select Case cn.Errors(0).NativeError
''        Case 2627
''            MsgBox "Trying with duplicate CNF Name"
''            txtName = ""
''            txtName.SetFocus
''        Case Else
''            MsgBox Err.Number & " : " & Err.Description
''    End Select
'
'End Function
'Public Sub FindRecord()
'If Not rsfactory.EOF Then
'        txtCID = rsfactory("sID")
''        txtCName = rsfactory("DCard")
'        txtCName = rsfactory("sName")
'        txtCAddress = rsfactory("sAddress")
'        txtCPhone = rsfactory("sPhone") & ""
'        txtCFax = IIf(IsNull(rsfactory("sProvince")), "", rsfactory("sProvince"))
'        txtCEmail = IIf(IsNull(rsfactory("sEmail")), "", rsfactory("sEmail"))
''        chkActive.Value = rsfactory("Active")
'    End If
'End Sub
'
'Private Function IsValidRecord() As Boolean
'    IsValidRecord = True
'    If (txtCName.text = "") Then
'       MsgBox "Enter Guest Name"
'       txtCName.SetFocus
'       IsValidRecord = False
'       Exit Function
'    End If
'
'    If (txtCAddress.text = "") Then
'      MsgBox "Enter Guest Address"
'      txtCAddress.SetFocus
'      IsValidRecord = False
'      Exit Function
'
'    End If
'
'    If (txtCPhone.text = "") Then
'      MsgBox "Enter Guest Contact No"
'      txtCPhone.SetFocus
'      IsValidRecord = False
'      Exit Function
'    End If
'
'If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
'        If rsfactory.RecordCount > 0 Then
'        If rsfactory.State <> 0 Then rsfactory.Close
'            rsfactory.Open "select * from SSuplierName where upper(DCard)='" & Strings.UCase(Strings.Trim(parseQuotes(txtDiscountCard))) & "'", cn
'
'             If Not rsfactory.EOF Then
'        MsgBox "This Card No already exists Please Enter Another.", vbInformation, Me.Caption & " - " & App.Title
'          txtDiscountCard.SetFocus
'          IsValidRecord = False
'         Exit Function
'            End If
'
'         End If
'        End If
'    End Function
''.............................................................................
'
'Public Sub printReport()
''On Error GoTo ErrorHan
'Dim strPath         As String
'Dim rsFactProf      As ADODB.Recordset
'Dim strSQL          As String
'
'
'    strPath = App.Path + "\reports\CustomerInformationPreview.rpt"
'
'    Set oReportApp = CreateObject("Crystal.CRPE.Application")
'    Set oReport = oReportApp.OpenReport(strPath)
'    Set oReportDatabase = oReport.Database
'    Set oReportDatabaseTables = oReportDatabase.Tables
'    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
'    Set ObjPrinterSetting = oReport.PrintWindowOptions
'
'
'    Set rsFactProf = New ADODB.Recordset
'If rsFactProf.State <> 0 Then rsFactProf.Close
'
'    strSQL = "select SSuplierName.sID,SSuplierName.CName,SSuplierName.CAddress, " & _
'             "  " & _
'             "SSuplierName.CPhone,SSuplierName.CFax,SSuplierName.CEmail " & _
'             "from SSuplierName where " & _
'             "SSuplierName.sID='" & Me.txtCID & "'"
'
'    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly
'
'    oReportDatabaseTable.SetPrivateData 3, rsFactProf
'
'ObjPrinterSetting.HasPrintSetupButton = True
'ObjPrinterSetting.HasRefreshButton = True
'ObjPrinterSetting.HasSearchButton = True
'ObjPrinterSetting.HasZoomControl = True
'
''      Set oReportFormulaFieldDefinations = oReport.FormulaFields
''      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
''      oReportFF.text = "'Factory Information'"
'
'oReport.DiscardSavedData
'oReport.Preview "Customer Infromation of '" & txtCName.text & "'", , , , , 16777216 Or 524288 Or 65536
'
'End Sub
'
'Public Sub PopulateCnf(StrID As String)
'    rsfactory.MoveFirst
'    rsfactory.Find "sID=" & parseQuotes(StrID)
'    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord
'
'End Sub
'
'
'
'
'
'
