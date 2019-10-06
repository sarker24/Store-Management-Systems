VERSION 5.00
Begin VB.Form frmSuppliername 
   BackColor       =   &H00C0B4A9&
   Caption         =   " Supplier Name"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   Icon            =   "frmSname.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   6495
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSerial 
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   465
      Left            =   1800
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   18
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtProvince 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Please Enter Only Numeric Number"
      Top             =   4200
      Width           =   4455
   End
   Begin VB.TextBox txtEmail 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   4
      Text            =   " "
      Top             =   4800
      Width           =   4455
   End
   Begin VB.TextBox txtphone 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Please Enter Only Numeric Number"
      Top             =   3600
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   3480
      Picture         =   "frmSname.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5400
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   795
      Left            =   4440
      Picture         =   "frmSname.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5400
      Width           =   945
   End
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   600
      Picture         =   "frmSname.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5400
      Width           =   945
   End
   Begin VB.CommandButton CmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1560
      Picture         =   "frmSname.frx":1FE8
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   945
   End
   Begin VB.CommandButton CmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   5400
      Picture         =   "frmSname.frx":28B2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
      Width           =   945
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete"
      Height          =   795
      Left            =   2520
      Picture         =   "frmSname.frx":317C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   945
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmSname.frx":381C
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Text            =   " "
      Top             =   1200
      Width           =   4455
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
      Left            =   240
      TabIndex        =   17
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      Caption         =   "Store Management System (SMS)"
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
      TabIndex        =   16
      Top             =   0
      Width           =   6495
   End
   Begin VB.Label lblCountry 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Province Name"
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
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblPhone 
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
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label lblAddress 
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
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblName 
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
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblID 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Supplier ID "
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
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "frmSuppliername"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsfactory             As ADODB.Recordset
Private strFileName           As String
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim str As String
'--------------------------------------------------------------
Private oReportApp                        As CRPEAuto.Application
Private oReport                           As CRPEAuto.Report
Private oReportDatabase                   As CRPEAuto.Database
Private oReportDatabaseTables             As CRPEAuto.DatabaseTables
Private oReportDatabaseTable              As CRPEAuto.DatabaseTable
'Private oReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private oReportFF                         As CRPEAuto.FormulaFieldDefinition
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions

Private Sub chameleonButton1_Click()
    Call printReport
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdNew.Caption = "&New"
    cmdEdit.Caption = "&Edit"
    CmdDelete.Enabled = True
    cmdOpen.Enabled = True
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    txtSerial.Enabled = False
    Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If idelete = vbYes Then
            cn.Execute "Delete From SSuplierName Where sID ='" & parseQuotes(txtSerial) & "'"
            Call allClear
            Refresh
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

'
Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdDelete.Enabled = False
        cmdOpen.Enabled = False
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(sID),0) as SerialNo from SSuplierName"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerial.text = Val(rs!SerialNo) + 1
            
        Call allenable
        txtName.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSerial.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
                cmdOpen.Enabled = True
'                chameleonButton1.Enabled = True
                Call alldisable
                s = txtName
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "sName='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
'
    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdDelete.Enabled = False
        cmdOpen.Enabled = False
ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                rsfactory.Requery

                Dim s As String
                s = txtName
                rsfactory.Find "sName='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    frmSupplierSearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from SSuplierName", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
   
    If Not rsfactory.EOF Then FindRecord
    
    txtSerial.Enabled = False
    
End Sub

Private Sub allenable()
    txtSerial.Enabled = True
    txtName.Enabled = True
    txtAddress.Enabled = True
    txtphone.Enabled = True
    txtProvince.Enabled = True
    txtEmail.Enabled = True
End Sub

Private Sub alldisable()
    txtSerial.Enabled = False
    txtName.Enabled = False
    txtAddress.Enabled = False
    txtphone.Enabled = False
    txtProvince.Enabled = False
    txtEmail.Enabled = False
End Sub

Private Sub allClear()
    txtSerial.text = ""
    txtName.text = ""
    txtAddress.text = ""
    txtphone.text = ""
    txtProvince.text = ""
    txtEmail.text = ""
End Sub

Private Function rcupdate() As Boolean
'    On Error GoTo ErrHandler
    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then
    
    cn.Execute "INSERT INTO SSuplierName(sID,sName,sAddress, " & _
                   " sPhone,sProvince,sEmail) " & _
                   " VALUES ('" & parseQuotes(txtSerial) & "','" & parseQuotes(txtName) & "','" & parseQuotes(txtAddress) & "', " & _
                   " '" & parseQuotes(txtphone) & "', " & _
                   " '" & parseQuotes(txtProvince) & "', " & _
                   " '" & parseQuotes(txtEmail) & "') "


          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else
    
    
cn.Execute "Update SSuplierName Set sName='" & parseQuotes(txtName) & _
                  "',sAddress='" & parseQuotes(txtAddress) & "',sPhone='" & parseQuotes(txtphone) & "', " & _
                  " sProvince='" & parseQuotes(txtProvince) & _
                  "',sEmail='" & parseQuotes(txtEmail) & "' " & _
                  " Where sID ='" & parseQuotes(txtSerial) & "' "

        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
'    Exit Sub
    Exit Function

End Function
Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtSerial = rsfactory("sID")
        txtName = rsfactory("sName")
        txtAddress = rsfactory("sAddress")
        txtphone = rsfactory("sPhone") & ""
        txtProvince = IIf(IsNull(rsfactory("sProvince")), "", rsfactory("sProvince"))
        txtEmail = IIf(IsNull(rsfactory("sEmail")), "", rsfactory("sEmail"))
    End If
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If (txtName.text = "") Then
       MsgBox "Enter Supplier Name"
       txtName.SetFocus
       IsValidRecord = False
       Exit Function
    End If
    
    If (txtAddress.text = "") Then
      MsgBox "Enter Supplier Address"
      txtAddress.SetFocus
      IsValidRecord = False
      Exit Function
    End If
    
If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
        If rsfactory.RecordCount > 0 Then
        If rsfactory.State <> 0 Then rsfactory.Close
            rsfactory.Open "select * from SSuplierName where upper(sName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtName))) & "'", cn

             If Not rsfactory.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
        End If
    End Function
'.............................................................................

Public Sub printReport()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf      As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\CustomerInformationPreview.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select SSuplierName.sID,SSuplierName.sName,SSuplierName.sAddress, " & _
             "  " & _
             "SSuplierName.sPhone,SSuplierName.sProvince,SSuplierName.sEmail " & _
             "from SSuplierName where " & _
             "SSuplierName.sID='" & Me.txtSerial & "'"

    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True

'      Set oReportFormulaFieldDefinations = oReport.FormulaFields
'      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
'      oReportFF.text = "'Factory Information'"

oReport.DiscardSavedData
oReport.Preview "Customer Infromation of '" & txtName.text & "'", , , , , 16777216 Or 524288 Or 65536

End Sub

Public Sub PopulateCnf(StrID As String)
    rsfactory.MoveFirst
    rsfactory.Find "sID=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub txtphone_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
    End Sub
    
Private Sub txtProvince_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
    End Sub
    
Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
    End Sub

    
