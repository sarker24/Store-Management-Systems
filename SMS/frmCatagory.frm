VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCatagory 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Product Catagory"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   Icon            =   "frmCatagory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Find Last"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Find Next"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Find Previous"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Find First"
      Top             =   2040
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   6615
      Begin VB.TextBox txtSerial 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   4215
      End
      Begin VB.TextBox txtCatagory 
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
         Left            =   2040
         TabIndex        =   8
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label lnlSerial 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Serial No"
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCatagory 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Product Catagory"
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
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   3960
      Picture         =   "frmCatagory.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   795
      Left            =   4920
      Picture         =   "frmCatagory.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   1080
      Picture         =   "frmCatagory.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton CmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   2040
      Picture         =   "frmCatagory.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton CmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   5880
      Picture         =   "frmCatagory.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete"
      Height          =   795
      Left            =   3000
      Picture         =   "frmCatagory.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   945
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1320
      Top             =   2040
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
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmCatagory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private rs                     As ADODB.Recordset
Private rsSCatagory              As ADODB.Recordset
'Private strStream             As ADODB.Stream
Private strFileName            As String
Private bRecordExists          As Boolean
Dim str                        As String
'Private rm                    As New ADODB.Recordset
'Private rc                    As New ADODB.Recordset

Private Sub cmdCancel_Click()

   cmdCancel.Enabled = False
   CmdNew.Enabled = True
    CmdEdit.Caption = "&Edit"
    CmdNew.Caption = "&New"
    cmdClose.Enabled = True
    CmdEdit.Enabled = True
    CmdOpen.Enabled = True
    txtSerial.Enabled = False
    Call allClear
'    txtCompanyID.Enabled = False
    Call alldisable
    If Not rsSCatagory.EOF Then FindRecord
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If idelete = vbYes Then
            cn.Execute "Delete From SCatagory Where SCID ='" & parseQuotes(txtSerial) & "'"
'            cn.Execute "DELETE FROM SSalesDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
            Call allClear
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

    txtSerial = Adodc1.Recordset!SCID
    txtCatagory = Adodc1.Recordset!SCName
    
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

    txtSerial = Adodc1.Recordset!SCID
    txtCatagory = Adodc1.Recordset!SCName


End If
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

    txtSerial = Adodc1.Recordset!SCID
    txtCatagory = Adodc1.Recordset!SCName
    
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

    txtSerial = Adodc1.Recordset!SCID
    txtCatagory = Adodc1.Recordset!SCName


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
'        txtSerial.Enabled = True
        Call allClear
'       ModFunction.TextEnable Me, True
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SCID),0) as SerialNo from SCatagory"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerial.text = Val(rs!SerialNo) + 1
           Call allenable
'           Call alldisable
        txtCatagory.SetFocus

    ElseIf CmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then

                txtSerial.Enabled = False
                CmdNew.Caption = "&New"
                CmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdOpen.Enabled = True
'                ModFunction.TextEnable Me, False
                Call alldisable
                s = txtCatagory
                rsSCatagory.Requery
                rsSCatagory.MoveFirst
                rsSCatagory.Find "SCName='" & parseQuotes(s) & "'"
               
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
        txtCatagory.SetFocus
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
                rsSCatagory.Requery

                Dim s As String
                s = txtCatagory
                rsSCatagory.Find "SCName='" & parseQuotes(s) & "'"
                
                FindRecord
            End If
        End If
    End If
End Sub


Private Sub cmdOpen_Click()
   strCallingForm = LCase("frmCatagory")
    frmCatagorySearch.Show vbModal
    CmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
    ModFunction.StartUpPosition Me
    Set rsSCatagory = New ADODB.Recordset
'    Set rsImage = New ADODB.Recordset
    rsSCatagory.Open "select  DISTINCT * from SCatagory", cn, adOpenStatic, adLockReadOnly
    
ModFunction.TextEnable Me, False
    
    Call alldisable

   If rsSCatagory.RecordCount > 0 Then

        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
   If Not rsSCatagory.EOF Then FindRecord
    
    txtSerial.Enabled = False
    
Adodc1.ConnectionString = "Driver={SQL Server};" & _
   "Server=" & sServerName & ";" & _
   "Database=" & SDatabaseName & ";" & _
   "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "SCatagory"

  Adodc1.Refresh
    
End Sub

Private Sub allClear()
txtSerial.text = ""
txtCatagory.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If CmdNew.Caption = "&Save" Then
      
        cn.Execute "INSERT INTO SCatagory(SCID,SCName) " & _
                   " VALUES ('" & parseQuotes(txtSerial) & "','" & parseQuotes(txtCatagory) & "')"
                   

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else

        cn.Execute "Update SCatagory Set SCName='" & parseQuotes(txtCatagory) & _
                  "'WHERE  SCID ='" & parseQuotes(txtSerial) & "' "
            
 
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rsSCatagory.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate SCatagory Name"
            txtCatagory = ""
            txtCatagory.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsSCatagory.EOF Then
        txtSerial = rsSCatagory("SCID")
        txtCatagory = rsSCatagory("SCName")
        
   End If
End Sub


Private Sub allenable()
    txtCatagory.Enabled = True
    
End Sub

Private Sub alldisable()
    txtCatagory.Enabled = False
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (txtCatagory.text = "") Then
       MsgBox "Enter SCatagory Name"
       txtCatagory.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtSerial.text = "") Then
     MsgBox "Enter SCID"
     txtSerial.SetFocus
     IsValidRecord = False
     Exit Function
     
    End If
    
    If CmdEdit.Caption <> "&Update" Or CmdEdit.Caption = "&Update" Then
        If rsSCatagory.RecordCount > 0 Then
        If rsSCatagory.State <> 0 Then rsSCatagory.Close
            rsSCatagory.Open "select * from SCatagory where upper(SCName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtCatagory))) & "'", cn

             If Not rsSCatagory.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtCatagory.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If

End Function

Public Sub PopulateSCatagory(StrID As String)


    rsSCatagory.MoveFirst
    rsSCatagory.Find "SCID=" & parseQuotes(StrID)
    If rsSCatagory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub




'Option Explicit
'
'Private rsfactory             As ADODB.Recordset
'Private strFileName           As String
'Private bRecordExists         As Boolean
'Private rm                    As New ADODB.Recordset
'Private rs                    As New ADODB.Recordset
'Dim str As String
'
'Private Sub cmdClose_Click()
'    Unload Me
'End Sub
'Private Sub cmdCancel_Click()
'    cmdCancel.Enabled = False
'    cmdNew.Enabled = True
'    cmdNew.Caption = "&New"
'    cmdEdit.Caption = "&Edit"
''    cmdDelete.Enabled = True
'    cmdOpen.Enabled = True
'    cmdClose.Enabled = True
'    cmdEdit.Enabled = True
'    Call allClear
'    Call alldisable
'End Sub
'
'Private Sub cmdDelete_Click()
' On Error GoTo ErrHandler
'     Dim idelete As Integer
'    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
'
'    If frmLogin.txtUName.text = "Admin" Then
'    If idelete = vbYes Then
'            cn.Execute "Delete From SCatagory Where SCID ='" & parseQuotes(txtSerial) & "'"
'            Call allClear
'    End If
'    End If
'ErrHandler:
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
'     End Select
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
''        cmdDelete.Enabled = False
'        cmdOpen.Enabled = False
'        Call allClear
'
'If rs.State <> 0 Then rs.Close
'        Call allenable
''        SCName.SetFocus
'    ElseIf cmdNew.Caption = "&Save" Then
'        Dim s As String
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdNew.Caption = "&New"
'                cmdEdit.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
''                cmdDelete.Enabled = True
'                cmdOpen.Enabled = True
'                Call alldisable
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
'        txtCatagory.SetFocus
'        cmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
''        cmdDelete.Enabled = False
'        cmdOpen.Enabled = False
'ElseIf cmdEdit.Caption = "&Update" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdEdit.Caption = "&Edit"
'                cmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
''                cmdDelete.Enabled = True
'                cmdOpen.Enabled = True
'                Call alldisable
'                rsfactory.Requery
'
'                Dim s As String
'
'            End If
'        End If
'    End If
'End Sub
'
'Private Sub cmdOpen_Click()
'    frmCatagorySearch.Show vbModal
'    cmdOpen.Enabled = True
'    cmdCancel.Enabled = True
'End Sub
'
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    On Error Resume Next
'   If (KeyCode = 13 And Me.ActiveControl.Name <> "SCName") Then SendKeys "{TAB}", True
'End Sub
'
'Private Sub Form_Load()
'
'    Call Connect
'       ModFunction.StartUpPosition Me
'
'
'    Set rsfactory = New ADODB.Recordset
'    rsfactory.Open "select * from SCatagory", cn, adOpenStatic, adLockReadOnly
'    Call alldisable
'   If rsfactory.RecordCount > 0 Then
'        bRecordExists = True
'    Else
'        bRecordExists = False
'    End If
'    If Not rsfactory.EOF Then FindRecord
'    txtSerial.Enabled = True
'End Sub
'
'Public Sub FindRecord()
'If Not rsfactory.EOF Then
'        txtSerial = rsfactory("SCID")
'        txtCatagory = rsfactory("SCName")
'    End If
'End Sub
'
'Private Sub allenable()
'    txtCatagory.Enabled = True
'
'End Sub
'
'Private Sub alldisable()
'    txtCatagory.Enabled = False
'
'End Sub
'
'Private Sub allClear()
'    txtCatagory.text = ""
'
'End Sub
'
''Private Sub CmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'''CmdNew.Caption = "CLICK HERE TO ADD NEW ITEM "
''CmdNew.FontBold = True
'''CmdNew.ForeColor = &H4080
''End Sub
'
'
'Private Function rcupdate() As Boolean
'    On Error GoTo ErrHandler
'    cn.BeginTrans
'    If cmdNew.Caption = "&Save" Then
'
'    cn.Execute "INSERT INTO SCatagory(SCName) " & _
'               " VALUES ('" & parseQuotes(txtCatagory) & "')"
'
'              rcupdate = True
'          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
'    Else
'
'
' cn.Execute "Update SCatagory Set SCName='" & parseQuotes(txtCatagory) & "'  " & _
'            " Where SCID ='" & parseQuotes(txtSerial) & "' "
'
'        rcupdate = True
'        MsgBox "Record Updated successfully", vbInformation, "Confirmation"
'    End If
'
'    cn.CommitTrans
'
'    Exit Function
'
'
'ErrHandler:
'    cn.RollbackTrans
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "Please select record first for update", vbInformation, "Confirmation"
'            cmdEdit.Caption = "&Edit"
'   End Select
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
'
'Private Function IsValidRecord() As Boolean
'    IsValidRecord = True
'    If (txtCatagory.text = "") Then
'       MsgBox "Enter Catagory Name"
'       txtCatagory.SetFocus
'       IsValidRecord = False
'       Exit Function
'    End If
'
'
'If cmdEdit.Caption = "&Update" Or cmdNew.Caption = "&Save" Then
'Set rsfactory = New ADODB.Recordset
''        If rsfactory.RecordCount > 0 Then
'        If rsfactory.State <> 0 Then rsfactory.Close
'            rsfactory.Open "select * from SCatagory where upper(SCName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtCatagory))) & "'", cn
'
'             If Not rsfactory.EOF Then
'        MsgBox "This Catagory already exists Please Enter Another Catagory Name.", vbInformation, Me.Caption & " - " & App.Title
'          txtCatagory.SetFocus
'          IsValidRecord = False
'         Exit Function
'            End If
'
'         End If
''        End If
'    End Function
'
'Public Sub PopulateCnf(StrID As String)
'rsfactory.Close
'rsfactory.Open "select * from SCatagory", cn, adOpenStatic, adLockReadOnly
'    rsfactory.MoveFirst
'    rsfactory.Find "SCID=" & parseQuotes(StrID)
'    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord
'End Sub
'
'
'
'
