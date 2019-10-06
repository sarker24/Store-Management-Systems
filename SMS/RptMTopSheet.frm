VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptMTopSheet 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Monthly Top Sheet"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   Icon            =   "RptMTopSheet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6585
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
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   720
         Top             =   1320
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
               Picture         =   "RptMTopSheet.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptMTopSheet.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptMTopSheet.frx":11C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   2160
         TabIndex        =   1
         Top             =   1200
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
         Left            =   1680
         TabIndex        =   2
         Top             =   480
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65142787
         CurrentDate     =   40571
      End
      Begin MSComCtl2.DTPicker SSTDate 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   480
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
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   65142787
         CurrentDate     =   41071
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
         Left            =   1680
         TabIndex        =   5
         Top             =   240
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
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Monthly Top Sheet Report"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "RptMTopSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim intempid As Integer
Dim str1 As String
Dim st As String
Dim StrTableM As String
Dim PurchaseDetail As String
Dim temp As String
Dim TempA As String
Dim temp1 As String
Dim Temp11 As String
Dim temp2 As String
Dim Temp22 As String
Dim Profit As String


'----------------------------

Private objReportApp                           As CRPEAuto.Application
Private objReport                              As CRPEAuto.Report
'Private objReport                             As CRPEAuto.Report
Private objReportDatabase                      As CRPEAuto.Database
Private objReportDatabaseTables                As CRPEAuto.DatabaseTables
Private objReportDatabaseTable                 As CRPEAuto.DatabaseTable
Private objReportDatabaseTableSub1             As CRPEAuto.DatabaseTable
Private objReportDatabaseTableSub2             As CRPEAuto.DatabaseTable
Private ObjPrinterSetting                      As CRPEAuto.PrintWindowOptions
Private objReportFFDs                          As CRPEAuto.FormulaFieldDefinitions
Private objReportFFD                           As CRPEAuto.FormulaFieldDefinition
Private strEmp                                 As String
Private rsEmp                                  As ADODB.Recordset
Private rsProfitSub1                           As ADODB.Recordset
Private rsProfitSub2                           As ADODB.Recordset
Private Tracer                                 As Integer


Private objReportFormulaFieldDefinations       As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                            As CRPEAuto.FormulaFieldDefinition


Private objReportSub                            As CRPEAuto.Report 'sub
Private objReportDatabaseSub                    As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub              As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub               As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub     As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                          As CRPEAuto.FormulaFieldDefinition
    
    
Private Sub FetchData()
        
        Set rsEmp = New ADODB.Recordset
       
            rsEmp.CursorLocation = adUseClient
       
'       rsEmp.Open "exec Stock_Status2 '" & SSFDate & "', '" & SSTDate & "'", cn, adOpenStatic, adLockReadOnly
        rsEmp.Open "exec Monthly_Top_Sheet '" & SSFDate.Value & "', '" & SSTDate.Value & "'", cn, adOpenStatic, adLockReadOnly

        
End Sub

Private Sub previewReport()
    Dim strPath As String
'    On Error GoTo errh
    If (rsEmp.RecordCount = 0) Then
    MsgBox "Data Not Abailable", vbInformation, "Confarmation"
    Exit Sub
    End If
    

    
    
    strPath = App.Path + "\reports\Monthly Top Sheet.rpt"
    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(strPath)



    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions
    
    
    
    Set objReportFormulaFieldDefinations = objReport.FormulaFields
'    Set objReportFF = objReportFormulaFieldDefinations.Item(5)
'            objReportFF.text = "'" + str(Profit) + "'"


    Set objReportFF = objReportFormulaFieldDefinations.Item(1)

              objReportFF.text = "'" + Format(SSFDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
             objReportFF.text = "'" + Format(SSTDate, "dd-MMM-yyyy") + "'"
             
             
             
             

       objReportDatabaseTable.SetPrivateData 3, rsEmp
          
        
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        objReport.DiscardSavedData

        objReport.Preview "Profit Information", , , , , 16777216 Or 524288 Or 65536
         If Tracer = 1 Then
         objReport.PrintOut
         End If
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub
'errh:
'If Err.Number = 20545 Then
'
'MsgBox "Request cancelled by the user", vbInformation, "Print"
'Err.Clear
'End If
'MsgBox Err.Description
Set rsEmp = Nothing


End Sub

Private Sub Form_Load()
'    Call Connect
    ModFunction.StartUpPosition Me
'    SSCurrentDate.Value = Date
        SSTDate.Value = Date
        SSFDate.Value = Date


    End Sub

Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "Preview"
            If Validate Then
                Tracer = 0
                Call FetchData
                Call previewReport
               End If
     Case "Print"
            If Validate Then
                Tracer = 1
                Call FetchData
                Call previewReport
               End If
     Case "Close"
               Unload Me
    End Select
End Sub
Private Function Validate() As Boolean
           Validate = True
        If SSFDate.Value > SSTDate.Value Then
            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
            SSFDate.SetFocus
            Validate = False
            Exit Function
        End If
    End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function


