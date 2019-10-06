VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form RptCustomer 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Department Report Statement"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   Icon            =   "RptCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDateSelect 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Date"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   6495
      Begin VB.OptionButton OpCustomDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&stom Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton OpCurrentDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&rrent Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1680
         Top             =   1920
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
               Picture         =   "RptCustomer.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptCustomer.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptCustomer.frx":11C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
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
      Begin MSComCtl2.DTPicker SSFDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16580611
         CurrentDate     =   38258
      End
      Begin MSComCtl2.DTPicker SSTDate 
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16580611
         CurrentDate     =   38258
      End
      Begin MSComCtl2.DTPicker SSCurrentDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16580611
         CurrentDate     =   38258
      End
      Begin MSAdodcLib.Adodc dcBudgetHead 
         Height          =   330
         Left            =   4200
         Top             =   1920
         Visible         =   0   'False
         Width           =   2160
         _ExtentX        =   3810
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
      Begin VB.Label lblFrom 
         BackColor       =   &H00C0B4A9&
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   4200
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   6495
      Begin VB.CheckBox chkAllItem 
         BackColor       =   &H00C0B4A9&
         Caption         =   "All Department"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin MSForms.ComboBox cmbBudgetHead 
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   2655
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4683;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Department  Name"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label lblIWSSales 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Department Report Information"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "RptCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsMaster                            As ADODB.Recordset
Private rsSelect                            As ADODB.Recordset 'sub

Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition
Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private Tracer                              As Integer
Private strGroupName                        As String

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub chkAllItem_Click()
If chkAllItem.Value = 1 Then
cmbBudgetHead.Enabled = False
cmbBudgetHead.text = ""
Else
cmbBudgetHead.Enabled = True
End If
End Sub

Private Sub Form_Load()
    Call Connect
    Call BudgetName
    ModFunction.StartUpPosition Me
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    SSCurrentDate.Value = Date
    SSTDate.Value = Date
    OpCurrentDate.Value = True
End Sub

Private Sub BudgetName()
    Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT sDName FROM SDeptName ORDER BY sDName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbBudgetHead.AddItem rsTemp2("sDName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub

Private Sub OpCurrentDate_Click()
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    SSCurrentDate.Visible = True
    lblFrom.Visible = False
    SSFDate.Visible = False
    lblTo.Visible = False
    SSTDate.Visible = False
End Sub
Private Sub opCustomDate_Click()
    OpCustomDate.Visible = True
    OpCurrentDate.Visible = True
    lblFrom.Visible = True
    lblTo.Visible = True
    SSCurrentDate.Visible = False
    SSFDate.Visible = True
    SSTDate.Visible = True
End Sub
Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "Preview"
            If Validate Then
            If chkAllItem.Value = 1 Then
                Tracer = 0
                    Call FetchData1
                        Call previewReport1
                    Else
                Call FetchData
            Call previewReport
                    End If
               End If
     Case "Print"
            If Validate Then
            If chkAllItem.Value = 1 Then
                Tracer = 0
                    Call FetchData1
                        Call previewReport1
                    Else
                Call FetchData
            Call previewReport
                    End If
               End If
     Case "Close"
               Unload Me
    End Select

End Sub
Private Function Validate() As Boolean
           Validate = True
        If SSFDate.Value > SSTDate.Value Then
            MsgBox "Invalid Date and select accurate date range", vbInformation, "Customer Wise Salse Report"
            SSFDate.SetFocus
            Validate = False
            Exit Function
        End If
    End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Public Function FetchData()

Set rsMaster = New ADODB.Recordset

    If OpCurrentDate.Value = True Then


   rsMaster.Open "SELECT SerialNo, ReqNo, DDate, DName, StoreName, Catagory, " & _
                 "ItemName, Qty, Rate,Qty *  Rate as Amt, Remarks, CPost, Unit " & _
                 "From SSalesDetails where  DDate='" & SSCurrentDate.Value & "' And CPost = 'Posted'and SSalesDetails.DName='" & parseQuotes(cmbBudgetHead) & "'", cn, adOpenStatic, adLockReadOnly


     End If

      If OpCustomDate.Value = True Then

rsMaster.Open "SELECT SerialNo, ReqNo, DDate, DName, StoreName, Catagory, " & _
              "ItemName, Qty, Rate,Qty *  Rate as Amt, Remarks, CPost, Unit " & _
              "From SSalesDetails where  DDate BETWEEN '" & SSFDate.Value & "' AND '" & SSTDate.Value & "' And CPost = 'Posted' and SSalesDetails.DName='" & parseQuotes(cmbBudgetHead) & "'", cn, adOpenStatic, adLockReadOnly

    End If


End Function

Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If


        strPath = App.Path + "\reports\Department Wise Reort.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields

   If OpCurrentDate.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
              objReportFF.text = "'" + Format(SSCurrentDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbBudgetHead + "'"

  End If


  If OpCustomDate.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)

              objReportFF.text = "'" + Format(SSFDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
             objReportFF.text = "'" + Format(SSTDate, "dd-MMM-yyyy") + "'"

             Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbBudgetHead + "'"

   End If



        objReportDatabaseTable.SetPrivateData 3, rsMaster

        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True

        objReport.DiscardSavedData
        objReport.Preview "Customer Wise Salse Statement", , , , , 16777216 Or 524288 Or 65536


     If Tracer = 1 Then
    objReport.PrintOut
    End If

        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Sales Summery Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Sales Summery Information Report"
    End Select
End Sub


Public Function FetchData1()

Set rsMaster = New ADODB.Recordset
    
    If OpCurrentDate.Value = True Then
    
               
rsMaster.Open "SELECT SerialNo, ReqNo, DDate, DName, StoreName,Catagory, " & _
                 "ItemName, Qty, Rate,Qty *  Rate as Amt, Remarks, CPost, Unit " & _
                 "From SSalesDetails where  DDate='" & SSCurrentDate.Value & "' And CPost = 'Posted'", cn, adOpenStatic, adLockReadOnly
                 
            
     End If
             
      If OpCustomDate.Value = True Then
      
                
rsMaster.Open "SELECT SerialNo, ReqNo, DDate, DName, StoreName,Catagory, " & _
                 "ItemName, Qty, Rate,Qty *  Rate as Amt, Remarks, CPost, Unit " & _
                 "From SSalesDetails where  DDate BETWEEN '" & SSFDate.Value & "' AND '" & SSTDate.Value & "' And CPost = 'Posted'", cn, adOpenStatic, adLockReadOnly
        
    End If
End Function

Public Sub previewReport1()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If


        strPath = App.Path + "\reports\Department Details Reort.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields

   If OpCurrentDate.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
              objReportFF.text = "'" + Format(SSCurrentDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbBudgetHead + "'"

  End If


  If OpCustomDate.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)

              objReportFF.text = "'" + Format(SSFDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
             objReportFF.text = "'" + Format(SSTDate, "dd-MMM-yyyy") + "'"

             Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbBudgetHead + "'"

   End If



        objReportDatabaseTable.SetPrivateData 3, rsMaster

        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True

        objReport.DiscardSavedData
        objReport.Preview "Customer Wise Salse Statement", , , , , 16777216 Or 524288 Or 65536


     If Tracer = 1 Then
    objReport.PrintOut
    End If

        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

ErrH:

    Select Case Err.Number
        Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Sales Summery Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Sales Summery Information Report"
    End Select
End Sub
