VERSION 5.00
Begin VB.Form frmDepartment 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Department Name"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   FillColor       =   &H00C0B4A9&
   Icon            =   "frmDepartment.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Delete"
      Height          =   795
      Left            =   2880
      Picture         =   "frmDepartment.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   945
   End
   Begin VB.CommandButton CmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   5760
      Picture         =   "frmDepartment.frx":0C2A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2640
      Width           =   945
   End
   Begin VB.CommandButton CmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1920
      Picture         =   "frmDepartment.frx":14F4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   945
   End
   Begin VB.CommandButton CmdNew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   960
      Picture         =   "frmDepartment.frx":1DBE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   795
      Left            =   4800
      Picture         =   "frmDepartment.frx":2688
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      Height          =   795
      Left            =   3840
      Picture         =   "frmDepartment.frx":2F52
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      Begin VB.TextBox txtDName 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   2040
         TabIndex        =   12
         Text            =   " "
         Top             =   1560
         Width           =   4215
      End
      Begin VB.TextBox txtSerial 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   495
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   " "
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtDCode 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   4215
      End
      Begin VB.Label lblDName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Deparment Name"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lblDID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Deparment ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblDCode 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Deparment Code"
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
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
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
      TabIndex        =   11
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private rs                     As ADODB.Recordset
Private rsSDeptName              As ADODB.Recordset
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
    cmdOpen.Enabled = True
    txtSerial.Enabled = False
    Call allClear
'    txtCompanyID.Enabled = False
    Call alldisable
    If Not rsSDeptName.EOF Then FindRecord
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If idelete = vbYes Then
            cn.Execute "Delete From SDeptName Where DID ='" & parseQuotes(txtSerial) & "'"
'            cn.Execute "DELETE FROM SSalesDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
            Call allClear
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
       Set rs = New ADODB.Recordset
    If CmdNew.Caption = "&New" Then
        CmdNew.Caption = "&Save"
        CmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False
'        txtSerial.Enabled = True
        Call allClear
'       ModFunction.TextEnable Me, True
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(sDID),0) as SerialNo from SDeptName"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerial.text = Val(rs!SerialNo) + 1
           Call allenable
'           Call alldisable
        txtDName.SetFocus

    ElseIf CmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then

                txtSerial.Enabled = False
                CmdNew.Caption = "&New"
                CmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
'                ModFunction.TextEnable Me, False
                Call alldisable
                s = txtDName
                rsSDeptName.Requery
                rsSDeptName.MoveFirst
                rsSDeptName.Find "sDName='" & parseQuotes(s) & "'"
               
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
        txtDCode.SetFocus
        CmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdOpen.Enabled = False

    ElseIf CmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                CmdEdit.Caption = "&Edit"
                CmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                rsSDeptName.Requery

                Dim s As String
                s = txtDName
                rsSDeptName.Find "sDName='" & parseQuotes(s) & "'"
                
                FindRecord
            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
   strCallingForm = LCase("frmSDeptName")
    frmDepartmentSearch.Show vbModal
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
    Set rsSDeptName = New ADODB.Recordset
'    Set rsImage = New ADODB.Recordset
    rsSDeptName.Open "select  DISTINCT * from SDeptName", cn, adOpenStatic, adLockReadOnly
    
ModFunction.TextEnable Me, False
    
    Call alldisable

   If rsSDeptName.RecordCount > 0 Then

        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
   If Not rsSDeptName.EOF Then FindRecord
    
    txtSerial.Enabled = False
    
End Sub

Private Sub allClear()
txtSerial.text = ""
txtDCode.text = ""
txtDName.text = ""
End Sub

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If CmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO SDeptName(sDID,sDCode,sDName) " & _
                   " VALUES ('" & parseQuotes(txtSerial) & "','" & parseQuotes(txtDCode) & "', " & _
                   " '" & parseQuotes(txtDName) & "')"
                   

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else

         cn.Execute "Update SDeptName Set sDCode='" & parseQuotes(txtDCode) & _
                  "',sDName='" & parseQuotes(txtDName) & "' where sDID='" & parseQuotes(txtSerial) & "'"
                  
                 
'        If (UCase(txtSerial.text) = UCase(rsSDeptName!sDID)) And UCase(txtDName.text) = UCase(rsSDeptName!SDeptNameName) Then
'    MsgBox "Trying Duplicate SDeptName Name"
'        Exit Function
'    End If
'
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rsSDeptName.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate SDeptName Name"
            txtDName = ""
            txtDName.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsSDeptName.EOF Then
        txtSerial = rsSDeptName("sDID")
        txtDCode = rsSDeptName("sDCode")
        txtDName = rsSDeptName("sDName")
        
   End If
End Sub


Private Sub allenable()
    txtDCode.Enabled = True
    txtDName.Enabled = True
    
End Sub

Private Sub alldisable()
    txtDCode.Enabled = False
    txtDName.Enabled = False
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True

   If (txtDCode.text = "") Then
       MsgBox "Enter SDeptName Code"
       txtDCode.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtDName.text = "") Then
       MsgBox "Enter SDeptName Name"
       txtDName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtSerial.text = "") Then
     MsgBox "Enter sDID"
     txtSerial.SetFocus
     IsValidRecord = False
     Exit Function
     
    End If
    
    If CmdEdit.Caption <> "&Update" Or CmdEdit.Caption = "&Update" Then
        If rsSDeptName.RecordCount > 0 Then
        If rsSDeptName.State <> 0 Then rsSDeptName.Close
            rsSDeptName.Open "select * from SDeptName where upper(sDName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtDName))) & "'", cn

             If Not rsSDeptName.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtDName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If

End Function

Public Sub PopulateSDeptName(StrID As String)


    rsSDeptName.MoveFirst
    rsSDeptName.Find "sDID=" & parseQuotes(StrID)
    If rsSDeptName.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub











