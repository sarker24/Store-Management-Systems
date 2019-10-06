VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCatagorySub 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Catagory Information"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   Icon            =   "frmSubCatagory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Find First"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Find Previous"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Find Next"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Find Last"
      Top             =   2760
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   8175
      Begin VB.TextBox txtIName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   5295
      End
      Begin VB.TextBox txtROL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6000
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtSerialNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   420
         Left            =   240
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0B4A9&
         Caption         =   "New..."
         Height          =   735
         Left            =   6960
         Picture         =   "frmSubCatagory.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin MSForms.ComboBox cmbStoreName 
         Height          =   375
         Left            =   1200
         TabIndex        =   24
         Top             =   480
         Width           =   2055
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3625;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblStoreName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Store Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin MSForms.ComboBox txtUnit 
         Height          =   420
         Left            =   5640
         TabIndex        =   23
         Top             =   1320
         Width           =   2415
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4260;741"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblItemName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Item Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   22
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label lblROL 
         BackColor       =   &H00C0B4A9&
         Caption         =   "ROL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6000
         TabIndex        =   21
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblUnit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5640
         TabIndex        =   20
         Top             =   1080
         Width           =   1935
      End
      Begin MSForms.ComboBox cmbItemCatagory 
         Height          =   375
         Left            =   3360
         TabIndex        =   0
         Top             =   480
         Width           =   2535
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "4471;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   975
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
         Height          =   375
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   795
      Left            =   3840
      Picture         =   "frmSubCatagory.frx":0FB4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   945
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete"
      Height          =   795
      Left            =   2040
      Picture         =   "frmSubCatagory.frx":187E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   945
   End
   Begin VB.CommandButton CmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1080
      Picture         =   "frmSubCatagory.frx":1F1E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   945
   End
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   120
      Picture         =   "frmSubCatagory.frx":27E8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3360
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   3000
      Picture         =   "frmSubCatagory.frx":30B2
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3360
      Width           =   945
   End
   Begin VB.CommandButton cmdpreview 
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
      Height          =   795
      Left            =   4800
      Picture         =   "frmSubCatagory.frx":397C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   945
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
      Height          =   795
      Left            =   5760
      Picture         =   "frmSubCatagory.frx":4246
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   945
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3360
      Top             =   2880
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DCSearch"
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
      Caption         =   "Inventory Management System (IMS)"
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
      TabIndex        =   10
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmCatagorySub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsSubCatagory        As ADODB.Recordset
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
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions

Private Sub chameleonButton1_Click()
'    Call printReport
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
'
Private Sub cmdCancel_Click()

    cmdCancel.Enabled = False
    CmdNew.Enabled = True
    CmdEdit.Caption = "&Edit"
    CmdNew.Caption = "&New"
    cmdClose.Enabled = True
    CmdEdit.Enabled = True
    CmdOpen.Enabled = True
'    chameleonButton1.Enabled = True
    txtSerialNo.Enabled = False
    Call allClear
    Call alldisable
    If Not rsSubCatagory.EOF Then FindRecord
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
     idelete = MsgBox("Do you want to delete this record?", vbYesNo)
     If frmLogin.txtUID.text = "Admin" Then
    If idelete = vbYes Then
  
    cn.Execute "Delete From SubCatagory Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
            Call allClear
    End If
        
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.EOF = True Then
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

    txtSerialNo = Adodc1.Recordset!SerialNo
    cmbStoreName = Adodc1.Recordset!StoreName
    cmbItemCatagory = Adodc1.Recordset!Catagory
    txtIName = Adodc1.Recordset!Itemname
    txtROL = Adodc1.Recordset!Rol
    txtUnit = Adodc1.Recordset!Unit

End If

End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

    txtSerialNo = Adodc1.Recordset!SerialNo
    cmbStoreName = Adodc1.Recordset!StoreName
    cmbItemCatagory = Adodc1.Recordset!Catagory
    txtIName = Adodc1.Recordset!Itemname
    txtROL = Adodc1.Recordset!Rol
    txtUnit = Adodc1.Recordset!Unit


End If
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If CmdNew.Caption = "&New" Then
        CmdNew.Caption = "&Save"
        CmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdOpen.Enabled = False
'        chameleonButton1.Enabled = False
        Call allClear
'        CALL
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),0) as SerialNo from SubCatagory"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo.text = Val(rs!SerialNo) + 1
            
        Call allenable
        cmbItemCatagory.SetFocus
    ElseIf CmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSerialNo.Enabled = False
                CmdNew.Caption = "&New"
                CmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdOpen.Enabled = True
                Call alldisable
                s = txtSerialNo
                rsSubCatagory.Requery
                rsSubCatagory.MoveFirst
                rsSubCatagory.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Sub cmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdNew.Enabled = False
        Call allenable
        cmbItemCatagory.SetFocus
        CmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        CmdOpen.Enabled = False
    ElseIf CmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                CmdEdit.Caption = "&Edit"
                CmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdOpen.Enabled = True
                Call alldisable
                rsSubCatagory.Requery

                Dim s As String
                s = txtSerialNo
                rsSubCatagory.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

    txtSerialNo = Adodc1.Recordset!SerialNo
    cmbStoreName = Adodc1.Recordset!StoreName
    cmbItemCatagory = Adodc1.Recordset!Catagory
    txtIName = Adodc1.Recordset!Itemname
    txtROL = Adodc1.Recordset!Rol
    txtUnit = Adodc1.Recordset!Unit


End If
End Sub

Private Sub cmdOpen_Click()
    frmCatagorySubItemSearch.Show vbModal
    CmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub Find_Click()
    CmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub cmdPreview_Click()
RptItemSearh.Show vbModal
End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
       cmdPrevious.Enabled = False
 Else
      cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

    txtSerialNo = Adodc1.Recordset!SerialNo
    cmbStoreName = Adodc1.Recordset!StoreName
    cmbItemCatagory = Adodc1.Recordset!Catagory
    txtIName = Adodc1.Recordset!Itemname
    txtROL = Adodc1.Recordset!Rol
    txtUnit = Adodc1.Recordset!Unit

End If
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsSubCatagory = New ADODB.Recordset
    rsSubCatagory.Open "select * from SubCatagory order by ItemName", cn, adOpenStatic, adLockReadOnly
    
    Call alldisable
    Call Itemname
    Call PStoreName
    
   If rsSubCatagory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If

    If Not rsSubCatagory.EOF Then FindRecord

'    txtSerialNo.Enabled = False

    txtUnit.AddItem "PCS"
    txtUnit.AddItem "PACK"
    txtUnit.AddItem "ROLL"
    txtUnit.AddItem "BOTTLE"
    txtUnit.AddItem "kIT"

 Adodc1.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "SubCatagory"

  Adodc1.Refresh
    
End Sub

Private Sub cmbStoreName_GotFocus()
cmbStoreName.SelStart = 0
cmbStoreName.SelLength = Len(cmbStoreName)

End Sub


Private Sub cmbItemCatagory_GotFocus()
cmbItemCatagory.SelStart = 0
cmbItemCatagory.SelLength = Len(cmbItemCatagory)

End Sub

Private Sub PStoreName()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT StoreName FROM HStore ORDER BY StoreName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbStoreName.AddItem rsTemp2("StoreName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub


Private Sub Itemname()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT SCName FROM SCatagory ORDER BY SCName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbItemCatagory.AddItem rsTemp2("SCName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub allenable()
    txtIName.Enabled = True
    cmbStoreName.Enabled = True
    cmbItemCatagory.Enabled = True
    txtUnit.Enabled = True
    txtROL.Enabled = True
    End Sub

Private Sub alldisable()
    txtSerialNo.Enabled = False
    cmbStoreName.Enabled = False
    txtIName.Enabled = False
    cmbItemCatagory.Enabled = False
    txtUnit.Enabled = False
    txtROL.Enabled = False
    
End Sub


Private Sub allClear()
    txtSerialNo.text = ""
    cmbStoreName.text = ""
    txtIName.text = ""
    cmbItemCatagory.text = ""
    txtUnit.text = ""
    txtROL.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If CmdNew.Caption = "&Save" Then
    
    cn.Execute "INSERT INTO SubCatagory( SerialNo, StoreName,Catagory, ItemName, ROL, Unit) " & _
               " VALUES ('" & parseQuotes(txtSerialNo) & "','" & parseQuotes(cmbStoreName) & "','" & parseQuotes(cmbItemCatagory) & "', " & _
               " '" & parseQuotes(txtIName) & "'," & parseQuotes(txtROL) & ",'" & parseQuotes(txtUnit) & "')"
                   
                   
    rcupdate = True
    MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    
    Else

    cn.Execute "Update SubCatagory Set StoreName='" & parseQuotes(cmbStoreName) & "',Catagory='" & parseQuotes(cmbItemCatagory) & "',ItemName='" & parseQuotes(txtIName) & _
               "',ROL='" & parseQuotes(txtROL) & "',Unit='" & parseQuotes(txtUnit) & "' WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"

                  
                 
     rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans

    Exit Function



ErrHandler:
    cn.RollbackTrans
    rsSubCatagory.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate SubCatagory Name"
            cmbItemCatagory = ""
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsSubCatagory.EOF Then
        txtSerialNo = rsSubCatagory("SerialNo")
        cmbStoreName = rsSubCatagory("StoreName")
        cmbItemCatagory = rsSubCatagory("Catagory")
        txtIName = rsSubCatagory("ItemName")
        txtROL = rsSubCatagory("ROL")
        txtUnit = rsSubCatagory("Unit")
End If
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If (txtIName.text = "") Then
       MsgBox "Enter Valid Item Name"
       txtIName.SetFocus
       IsValidRecord = False
       Exit Function
    End If
    If (cmbItemCatagory.text = "") Then
    MsgBox "Enter Item Catagory Name"
    cmbItemCatagory.SetFocus
    IsValidRecord = False
        Exit Function
    End If
    
    If (cmbStoreName.text = "") Then
    MsgBox "Enter Project Store Name"
    cmbStoreName.SetFocus
    IsValidRecord = False
        Exit Function
    End If
    
If CmdEdit.Caption <> "&Update" Or CmdEdit.Caption = "&Update" Then
        If rsSubCatagory.RecordCount > 0 Then
        If rsSubCatagory.State <> 0 Then rsSubCatagory.Close
            rsSubCatagory.Open "select * from SubCatagory where upper(ItemName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtIName))) & "'", cn

'        If Not rsSubCatagory.EOF Then
'        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
'          cmbItemCatagory.SetFocus
'          IsValidRecord = False
'         Exit Function
'            End If

         End If
    End If
End Function
'.............................................................................

Public Sub PopulateForm(StrID As String)

 Set rsSubCatagory = New ADODB.Recordset

    If rsSubCatagory.State <> 0 Then rsSubCatagory.Close
        rsSubCatagory.Open "select SerialNo,StoreName,Catagory,ItemName,ROL,Unit from SubCatagory", cn, adOpenStatic, adLockReadOnly
                          

        rsSubCatagory.MoveFirst
    
    rsSubCatagory.Find "SerialNo=" & parseQuotes(StrID)
    If rsSubCatagory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub



Public Sub printpreview()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf         As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\ItemPreview.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select SubCatagory.SerialNo,SubCatagory.StoreName,SubCatagory.Catagory, subCatagory.ItemName,SubCatagory.Unit"
             
    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True
oReport.DiscardSavedData
oReport.Preview "SubCatagory Infromation of '" & cmbItemCatagory.text & "'", , , , , 16777216 Or 524288 Or 65536


End Sub


Public Sub PopulateCnf(StrID As String)


    rsSubCatagory.MoveFirst
    rsSubCatagory.Find "SerialNo=" & parseQuotes(StrID)
    If rsSubCatagory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub




