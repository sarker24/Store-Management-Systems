VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmItemEntry 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Entry"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmItemEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLAdd 
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
      Left            =   9360
      Picture         =   "frmItemEntry.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Add"
      Top             =   960
      Width           =   300
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
      Left            =   9360
      Picture         =   "frmItemEntry.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Remove"
      Top             =   1320
      Width           =   300
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      Picture         =   "frmItemEntry.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   870
   End
   Begin VB.CommandButton chameleonButton1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      Picture         =   "frmItemEntry.frx":1968
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   990
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "C&lose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Picture         =   "frmItemEntry.frx":2232
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   870
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      Picture         =   "frmItemEntry.frx":2AFC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7080
      Width           =   990
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      Picture         =   "frmItemEntry.frx":33C6
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3720
      Picture         =   "frmItemEntry.frx":3C90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   870
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Picture         =   "frmItemEntry.frx":455A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   990
   End
   Begin VB.TextBox txtSerialNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   600
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "New..."
      Height          =   735
      Left            =   8880
      Picture         =   "frmItemEntry.frx":5224
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin VSFlex7LCtl.VSFlexGrid fgItem 
      Height          =   5850
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   9480
      _cx             =   16722
      _cy             =   10319
      _ConvInfo       =   1
      Appearance      =   2
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
      BackColor       =   12629161
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   12629161
      BackColorAlternate=   14737632
      GridColor       =   12629161
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmItemEntry.frx":590E
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbItemGroup 
      Height          =   400
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   2500
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
      _ExtentX        =   4410
      _ExtentY        =   706
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbItemCatagory 
      Height          =   400
      Left            =   6000
      TabIndex        =   2
      Top             =   360
      Width           =   2500
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
      _ExtentX        =   4410
      _ExtentY        =   706
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin MSAdodcLib.Adodc dcItemGroup 
      Height          =   360
      Left            =   7920
      Top             =   6960
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   635
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
   Begin MSAdodcLib.Adodc dcCatagory 
      Height          =   360
      Left            =   7920
      Top             =   7320
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   635
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
      Caption         =   "dcCatagory"
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
      BackColor       =   &H00C0B4A9&
      Caption         =   "Serial No"
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
      Left            =   600
      TabIndex        =   13
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblegroup 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Item Group"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblItemCatagory 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Item Catagory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmItemEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Dim rs As New ADODB.Recordset
 Private rsItemMaster                 As ADODB.Recordset
 Private rsItemDetail                 As ADODB.Recordset
 Private rs                              As ADODB.Recordset
' Set rsItemDetail = New ADODB.Recordset
' Private rsDyeingBuyer                 As ADODB.Recordset
' Private rsTemp                         As ADODB.Recordset
' Public rsDyeingMeasurement                     As ADODB.Recordset
''      ' Public rsMeasurement                  As ADODB.Recordset
''      ' Public rsYarn                         As ADODB.Recordset
 
 Private bRecordExists                  As Boolean
' Dim flagSlNo                           As Integer
 Dim str As String
' Dim s As String
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
Private temp As Double
Private temp1 As Double
'--------------------------------------------------------------------------------


Private Sub chameleonButton1_Click()
    Call printReport
End Sub


'Private Sub cmbItemCatagory_Change()
''dcEmpCategory_Change
'Call ItemCatagory
'End Sub

Private Sub cmdCancel_Click()
    
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    cmdDelete.Enabled = True
    chameleonButton1.Enabled = True
    Call alldisable
    If Not rsItemMaster.EOF Then FindRecord
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdEdit_Click()
     If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        cmbItemGroup.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
        chameleonButton1.Enabled = False
        cmdLAdd.Enabled = True
        cmdLDelete.Enabled = True
        fgItem.Editable = flexEDKbdMouse
        txtSerialNo.Enabled = False
        
    ElseIf cmdEdit.Caption = "&Update" Then
'          Call duplicate
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                chameleonButton1.Enabled = True
                cmdDelete.Enabled = True
                cmdClose.Enabled = True
                fgItem.Editable = flexEDNone
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

End Sub

Private Sub cmdLAdd_Click()
With fgItem
        If .Row = -1 Or .Row = 0 Then
            .AddItem ""
            Exit Sub
        End If
        If .Row > 0 Then
                .AddItem "", .Row + 1
        End If
    End With
    
End Sub

Private Sub cmdLDelete_Click()
            With fgItem
        If .Row = 0 Or .Row = -1 Then Exit Sub

        If .Rows > 1 Then .RemoveItem .Row
    End With

End Sub

Private Sub cmdNew_Click()
    
    Set rs = New ADODB.Recordset
If cmdNew.Caption = "&New" Then
        
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
        cmdOpen.Enabled = False
        cmdClose.Enabled = False
        cmdLAdd.Enabled = True
        cmdLDelete.Enabled = True
        chameleonButton1.Enabled = False
'        TextClear Me
        Call Clear
         
        fgItem.Rows = 1
        fgItem.Editable = flexEDKbdMouse
        Call allenable
        cmbItemGroup.SetFocus
           If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),1) as InvNo from tblItemMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo = Val(rs!InvNo) + 1

        
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
'        Call duplicate
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdDelete.Enabled = True
                cmdOpen.Enabled = True
                cmdCancel.Enabled = True
                chameleonButton1.Enabled = True
                
                Call alldisable
                s = txtSerialNo
                rsItemMaster.Requery
                rsItemMaster.MoveFirst
                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord
            End If
        End If
    End If
          
End Sub

Private Sub Clear()
    txtSerialNo.text = ""
'    DTPPayment.text = ""
    cmbItemGroup = ""
    cmbItemCatagory = ""
    
End Sub

Private Sub allenable()
    txtSerialNo.Enabled = True
    cmbItemGroup.Enabled = True
    cmbItemCatagory.Enabled = True
    fgItem.Enabled = True
    End Sub

Private Sub alldisable()
    txtSerialNo.Enabled = False
    cmbItemGroup.Enabled = False
    cmbItemCatagory.Enabled = False
    cmdLAdd.Enabled = False
    cmdLDelete.Enabled = False
    fgItem.Enabled = False

    
End Sub

Private Sub cmdOpen_Click()
    frmItemSearch.Show vbModal
        
End Sub
    
Private Sub Command1_Click()
'frmItemGroup.Show vbModal
End Sub


 Private Sub Form_Load()
         Call Connect
     ModFunction.StartUpPosition Me
       Call alldisable
       Call ItemGroup
        Call ItemCatagory
   Set rsItemMaster = New ADODB.Recordset
 
  If rsItemMaster.State <> 0 Then rsItemMaster.Close
     rsItemMaster.Open "select SerialNo,ItemGroup, " & _
                         "ItemCatagory FROM tblItemMaster order by SerialNo ", cn, adOpenStatic, adLockReadOnly

If rsItemMaster.RecordCount > 0 Then
      rsItemMaster.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
     
    If Not rsItemMaster.EOF Then FindRecord
    txtSerialNo.Enabled = False
    
End Sub

Private Sub ItemGroup()
    dcItemGroup.CursorLocation = adUseClient
    dcItemGroup.ConnectionString = cn.ConnectionString
    dcItemGroup.LockType = adLockReadOnly
    dcItemGroup.RecordSource = "SELECT distinct MenuGroup as ItemGroup FROM tbItemlGroupSetup ORDER BY ItemGroup"
    cmbItemGroup.DataMode = ssDataModeBound
    Set cmbItemGroup.DataSource = dcItemGroup
    cmbItemGroup.DataSourceList = dcItemGroup
    cmbItemGroup.DataFieldList = "ItemGroup"
    cmbItemGroup.DataField = "ItemGroup"
    cmbItemGroup.BackColorOdd = &HFFFF00
    cmbItemGroup.BackColorEven = &HFFC0C0
    cmbItemGroup.ForeColorEven = &H80000008
End Sub

Private Sub ItemCatagory()
    dcCatagory.CursorLocation = adUseClient
    dcCatagory.ConnectionString = cn.ConnectionString
    dcCatagory.LockType = adLockReadOnly
'    dcCatagory.RecordSource = "SELECT MenuCatagory, SerialNo as No, MenuGroup FROM tbItemlGroupSetup where MenuGroup='" & parseQuotes(cmbItemGroup.text) & "'"
    dcCatagory.RecordSource = "SELECT MenuCatagory, SerialNo as No, MenuGroup FROM tbItemlGroupSetup ORDER BY MenuCatagory"
    cmbItemCatagory.DataMode = ssDataModeBound
    Set cmbItemCatagory.DataSource = dcCatagory
    cmbItemCatagory.DataSourceList = dcCatagory
    cmbItemCatagory.DataFieldList = "MenuCatagory"
    cmbItemCatagory.DataField = "MenuCatagory"
    cmbItemCatagory.BackColorOdd = &HFFFF00
    cmbItemCatagory.BackColorEven = &HFFC0C0
    cmbItemCatagory.ForeColorEven = &H80000008
End Sub

 Private Function rcupdate() As Boolean

On Error GoTo ErrHandler
    Dim strSQL As String
    Dim iRow As Integer
    Dim j As Integer
'    flagSlNo = 0
   cn.BeginTrans
'     flagSlNo = 0
    If cmdNew.Caption = "&Save" Then
        
    'General Information for Payment Master
     strSQL = "INSERT INTO tblItemMaster (SerialNo, ItemGroup, ItemCatagory" & _
                ") " & _
                "VALUES ('" & txtSerialNo & "','" & cmbItemGroup.Columns(0).text & "', " & _
                " " & _
                "'" & cmbItemCatagory.Columns(0).text & "')"
        cn.Execute strSQL
    'payment Detail Information Enter This table
            j = 0
            For j = 1 To fgItem.Rows - 1
''                If fgExport.TextMatrix(j, 1) <> "" Then
                cn.Execute "INSERT INTO tblItemDetail (SerialNo,ItemGroup,ItemCatagory,ItemCode,ItemName,ItemQty,ItemPrice) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbItemGroup.Columns(0).text & "','" & cmbItemCatagory.Columns(0).text & "'," & _
                           "'" & parseQuotes(fgItem.TextMatrix(j, 4)) & "','" & parseQuotes(fgItem.TextMatrix(j, 5)) & "'," & _
                           "" & IIf(fgItem.TextMatrix(j, 6) = "", "0", fgItem.TextMatrix(j, 6)) & ", " & _
                           "" & IIf(fgItem.TextMatrix(j, 7) = "", "0", fgItem.TextMatrix(j, 7)) & ")"
                           Next
        rcupdate = True
        MsgBox "Record added Successfully", vbInformation, "Confirmation"
    Else
    ' Update Information
    
    
    

    
    
        cn.Execute "UPDATE tblItemMaster SET  ItemGroup = '" & cmbItemGroup.Columns(0).text & "', " & _
                 "ItemCatagory='" & cmbItemCatagory.Columns(0).text & "' " & _
                 " WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM tblItemDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


        j = 0
            For j = 1 To fgItem.Rows - 1
''                If fgExport.TextMatrix(j, 1) <> "" Then
                 cn.Execute "INSERT INTO tblItemDetail (SerialNo,ItemGroup,ItemCatagory,ItemCode,ItemName,ItemQty,ItemPrice) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbItemGroup.Columns(0).text & "','" & cmbItemCatagory.Columns(0).text & "'," & _
                           "'" & parseQuotes(fgItem.TextMatrix(j, 4)) & "','" & parseQuotes(fgItem.TextMatrix(j, 5)) & "'," & _
                           "" & IIf(fgItem.TextMatrix(j, 6) = "", "0", fgItem.TextMatrix(j, 6)) & ", " & _
                           "" & IIf(fgItem.TextMatrix(j, 7) = "", "0", fgItem.TextMatrix(j, 7)) & ")"
            Next

        rcupdate = True
        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
    End If
'    flagSlNo = 1
    cn.CommitTrans
    
    Exit Function
    
ErrHandler:
    cn.RollbackTrans
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Group Number"
            cmbItemGroup.SetFocus
        Case Else
'   If Err.Number = -2147217874 Then
'    MsgBox "You can't Insert same item from same style multiple times in one BTB LC."
''   End If
            MsgBox cn.Errors(0).NativeError & " : " & cn.Errors(0).Description
    End Select
End Function

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If Trim(cmbItemGroup) = "" Then
        MsgBox "Your are missing Menu Group Information", vbInformation
        cmbItemGroup.SetFocus
        IsValidRecord = False
        Exit Function
    ElseIf Trim(cmbItemCatagory) = "" Then
        MsgBox "Your are missing Iteam Catagory Information.", vbInformation
        cmbItemCatagory.SetFocus
        IsValidRecord = False
        Exit Function
        

'    ---------------------------------------------------

'ElseIf cmdNew.Caption = "&Save" Or cmdEdit.Caption = "&Update" Then
'         Dim k As Integer
'         If rsItemDetail.RecordCount > 0 Then
'         If rsItemDetail.State <> 0 Then rsItemDetail.Close
'            rsItemDetail.Open "select * from tblItemDetail where ItemCode='" & fgItem.TextMatrix(Row, 4) & "'", cn, adOpenStatic, adLockReadOnly
'
'             If Not rsItemDetail.EOF Then
'        MsgBox "This Record exists Duplicate ItemCode No.", vbInformation, Me.Caption & " - " & App.Title
'          fgItem.TextMatrix(k, 4).SetFocus
'          IsValidRecord = False
'         Exit Function
'            End If
'         End If
'         End If
'    Exit Function
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
    Else
        
        Dim j As Integer
        
         For j = 1 To fgItem.Rows - 2
        
        If Val(fgItem.TextMatrix(j, 4)) = Val(fgItem.TextMatrix(j + 1, 4)) Then
        MsgBox "Duplicate Item Code Number.", vbInformation
'         fgItem.TextMatrix(j, 4) = ""
'         fgItem.RemoveItem fgItem.Row
        IsValidRecord = False
        
        End If

       Next
       
       Exit Function
     End If
    End Function


Private Sub FindRecord()

    Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset
    txtSerialNo = rsItemMaster!SerialNo
    cmbItemGroup = rsItemMaster!ItemGroup
    cmbItemCatagory = rsItemMaster!ItemCatagory
    

    fgItem.Rows = 1
    strPaymentDetail = "SELECT  SerialNo, ItemGroup, ItemCatagory, ItemCode,ItemName,ItemQty,ItemPrice " & _
                " FROM tblItemDetail " & _
                "WHERE SerialNo='" & parseQuotes(txtSerialNo.text) & "' order by SerialNo "
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgItem.Rows = rsItemDetail.RecordCount + 1
                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgItem.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgItem.TextMatrix(i, 2) = rsItemDetail("ItemGroup")
            fgItem.TextMatrix(i, 3) = rsItemDetail("ItemCatagory")
            fgItem.TextMatrix(i, 4) = rsItemDetail("ItemCode")
            fgItem.TextMatrix(i, 5) = rsItemDetail("ItemName")
            fgItem.TextMatrix(i, 6) = rsItemDetail("ItemQty")
            fgItem.TextMatrix(i, 7) = rsItemDetail("ItemPrice")
            
       rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
End Sub


Public Sub printReport()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    If rsItemMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\MenuItemPreview.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        


    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close


            strSQL = "SELECT tblItemMaster.SerialNo, tblItemMaster.ItemGroup,tblItemMaster.ItemCatagory, " & _
                     "tblItemDetail.ItemCode, tblItemDetail.ItemName, tblItemDetail.ItemPrice,tblItemDetail.SerialNo " & _
                     "FROM tblItemDetail,tblItemMaster " & _
                     " WHERE tblItemMaster.SerialNo = tblItemDetail.SerialNo ORDER BY tblItemDetail.ItemCode"
                     
'                     " & _
                     "Where tblItemDetail.SerialNo='" & Me.txtSerialNo & "'"


'           strSQL = "SELECT USPaymentMaster.SerialNo,USPaymentMaster.PDate, USPaymentMaster.ReciveBy, " & _
'                    "USPaymentMaster.PaymentBy, USPaymentDetail.Description, " & _
'                    "USPaymentDetail.BillNo, USPaymentDetail.PaidTaka " & _
'                    " " & _
'                    "FROM USPaymentMaster INNER JOIN " & _
'                    "USPaymentDetail ON " & _
'                    "USPaymentMaster.SerialNo = USPaymentDetail.SerialNo " & _
'                    "Where USPaymentMaster.SerialNo ='" & Me.txtSerialNo & "'"

                
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
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Menu Item List Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Menu Item List Report"
    End Select
End Sub


Public Sub PopulateForm(StrID As String)
    rsItemMaster.MoveFirst
    rsItemMaster.Find "SerialNo=" & parseQuotes(StrID)
    If rsItemMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub



Private Sub duplicate()
Dim j As Integer
        
         For j = 1 To fgItem.Rows - 2
        
        If Val(fgItem.TextMatrix(j, 4)) = Val(fgItem.TextMatrix(j + 1, 4)) Then
        MsgBox "Duplicate Item Code Number.", vbInformation
         fgItem.TextMatrix(j, 4) = ""
         End If

Next

End Sub
Private Sub fgItem_AfterEdit(ByVal Row As Long, ByVal Col As Long)

'ElseIf cmdNew.Caption = "&Save" Or cmdEdit.Caption = "&Update" Then
         Dim k As Integer
         
         Set rsItemDetail = New ADODB.Recordset
'         If rsItemDetail.RecordCount > 0 Then
        
         If rsItemDetail.State <> 0 Then rsItemDetail.Close
         If Col = 4 Then
            rsItemDetail.Open "select * from tblItemDetail where ItemCode='" & fgItem.TextMatrix(Row, 4) & "'", cn, adOpenStatic, adLockReadOnly

             If Not rsItemDetail.EOF Then
        MsgBox "This Record exists Duplicate Item Code No.", vbInformation, Me.Caption & " - " & App.Title
'          fgItem.TextMatrix(Row, 4).
'          IsValidRecord = False
'         Exit Function
            End If
         End If
'         End If
'    Exit Function
End Sub



