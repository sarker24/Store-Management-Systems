VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmStore 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Store Information"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6930
   Icon            =   "frmStore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete"
      Height          =   795
      Left            =   3000
      Picture         =   "frmStore.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton CmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   5880
      Picture         =   "frmStore.frx":06AC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton CmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   2040
      Picture         =   "frmStore.frx":0F76
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   1080
      Picture         =   "frmStore.frx":1840
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   795
      Left            =   4920
      Picture         =   "frmStore.frx":210A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   3960
      Picture         =   "frmStore.frx":29D4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   6615
      Begin VB.TextBox txtStoreName 
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
         TabIndex        =   6
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtSID 
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
         TabIndex        =   5
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblCatagory 
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
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1815
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
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Find First"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Find Previous"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   375
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Find Next"
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   375
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Find Last"
      Top             =   2040
      Width           =   735
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
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
      TabIndex        =   15
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private rs                     As ADODB.Recordset
Private rsHStore              As ADODB.Recordset
Private strFileName            As String
Private bRecordExists          As Boolean
Dim str                        As String

Private Sub cmdCancel_Click()

   cmdCancel.Enabled = False
   CmdNew.Enabled = True
    CmdEdit.Caption = "&Edit"
    CmdNew.Caption = "&New"
    cmdClose.Enabled = True
    CmdEdit.Enabled = True
    CmdOpen.Enabled = True
    txtSID.Enabled = False
    Call allClear
    Call alldisable
    If Not rsHStore.EOF Then FindRecord
End Sub

Private Sub cmdClose_Click()
Unload Me
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

    txtSID = Adodc1.Recordset!SID
    txtStoreName = Adodc1.Recordset!StoreName
    
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

    txtSID = Adodc1.Recordset!SID
    txtStoreName = Adodc1.Recordset!StoreName


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
'        txtSID.Enabled = True
        Call allClear
'       ModFunction.TextEnable Me, True
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SID),0) as SerialNo from HStore"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSID.text = Val(rs!SerialNo) + 1
           Call allenable
'           Call alldisable
        txtStoreName.SetFocus

    ElseIf CmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then

                txtSID.Enabled = False
                CmdNew.Caption = "&New"
                CmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdOpen.Enabled = True
'                ModFunction.TextEnable Me, False
                Call alldisable
                s = txtStoreName
                rsHStore.Requery
                rsHStore.MoveFirst
                rsHStore.Find "StoreName='" & parseQuotes(s) & "'"
               
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
        txtStoreName.SetFocus
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
                rsHStore.Requery

                Dim s As String
                s = txtStoreName
                rsHStore.Find "StoreName='" & parseQuotes(s) & "'"
                
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

    txtSID = Adodc1.Recordset!SID
    txtStoreName = Adodc1.Recordset!StoreName


End If
End Sub

Private Sub cmdOpen_Click()
   strCallingForm = LCase("frmHStore")
    frmStoreSearch.Show vbModal
    CmdOpen.Enabled = True
    cmdCancel.Enabled = True
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

    txtSID = Adodc1.Recordset!SID
    txtStoreName = Adodc1.Recordset!StoreName
    
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
    ModFunction.StartUpPosition Me
    Set rsHStore = New ADODB.Recordset
    rsHStore.Open "select * from HStore", cn, adOpenStatic, adLockReadOnly

    
ModFunction.TextEnable Me, False
    
    Call alldisable

   If rsHStore.RecordCount > 0 Then

        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
   If Not rsHStore.EOF Then FindRecord
    
    txtSID.Enabled = False
    
Adodc1.ConnectionString = "Driver={SQL Server};" & _
       "Server=" & sServerName & ";" & _
       "Database=" & SDatabaseName & ";" & _
       "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "HStore"

  Adodc1.Refresh
    
End Sub

Private Sub allClear()
'    ModFunction.TextClear Me
txtStoreName.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If CmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO HStore(SID,StoreName) " & _
                   " VALUES ('" & parseQuotes(txtSID) & "','" & parseQuotes(txtStoreName) & "')"
                   

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else

        cn.Execute "Update HStore Set StoreName='" & parseQuotes(txtStoreName) & _
                  "'WHERE  SID ='" & parseQuotes(txtSID) & "' "
 
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rsStore.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Store Name"
            txtStoreName = ""
            txtStoreName.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsHStore.EOF Then
        txtSID = rsHStore("SID")
        txtStoreName = rsHStore("StoreName")
        
   End If
End Sub


Private Sub allenable()
    txtStoreName.Enabled = True
    
End Sub

Private Sub alldisable()
    txtStoreName.Enabled = False
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (txtStoreName.text = "") Then
       MsgBox "Enter Store Name"
       txtStoreName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtSID.text = "") Then
     MsgBox "Enter SID"
     txtSID.SetFocus
     IsValidRecord = False
     Exit Function
     
    End If
    
    If CmdEdit.Caption <> "&Update" Or CmdEdit.Caption = "&Update" Then
        If rsHStore.RecordCount > 0 Then
        If rsHStore.State <> 0 Then rsHStore.Close
            rsHStore.Open "select * from HStore where upper(StoreName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtStoreName))) & "'", cn

             If Not rsHStore.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtStoreName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
    
'    If cmdEdit.Caption <> "&Update" Then
'        If rsHStore.RecordCount > 0 Then
'            If (UCase(txtStoreName.text) = UCase(rsHStore!StoreName)) Then
'                  MsgBox "Trying Duplicate Store Name"
'                  IsValidRecord = False
'                 Exit Function
'            End If
'         End If
'    End If


End Function

Public Sub PopulateHStore(StrID As String)


    rsHStore.MoveFirst
    rsHStore.Find "SID=" & parseQuotes(StrID)
    If rsHStore.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


