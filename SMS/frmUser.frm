VERSION 5.00
Begin VB.Form frmUser 
   BackColor       =   &H00C0B4A9&
   Caption         =   " User Information"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "User Information Entry"
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
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5655
      Begin VB.CommandButton CmdDelete 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Delete"
         Height          =   735
         Left            =   2400
         Picture         =   "frmUser.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Width           =   1065
      End
      Begin VB.TextBox txtSerialNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   3135
      End
      Begin VB.ComboBox cboEn 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   8
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtConPas 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox TxtUName 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Q&uit"
         Height          =   735
         Left            =   3360
         Picture         =   "frmUser.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&New"
         Height          =   735
         Left            =   480
         Picture         =   "frmUser.frx":171E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Edit"
         Height          =   735
         Left            =   1440
         Picture         =   "frmUser.frx":1FE8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton cmdOpen 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Open"
         Height          =   735
         Left            =   4320
         Picture         =   "frmUser.frx":28B2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label lblSerial 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Confirm Passward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Password "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblPgroup 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Privilege Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
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
      Width           =   5775
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsfactory               As ADODB.Recordset
Private strFileName             As String
Private rsHouseName             As ADODB.Recordset
Private bRecordExists           As Boolean
Private rm                      As New ADODB.Recordset
Private rs                      As New ADODB.Recordset
Dim str As String

Private Sub cmdClose_Click()
    Unload Me
End Sub


Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    
    If frmLogin.txtUID.text = "Admin" Then
    If idelete = vbYes Then
            cn.Execute "Delete From SUser Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
            Call allClear
    End If
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

'
Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdExit.Enabled = False
        cmdOpen.Enabled = False
        CmdDelete.Enabled = False
        Call allClear

If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),0) as SerialNo from SUser"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo.text = Val(rs!SerialNo) + 1

        Call allenable
        txtUName.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
'                txtEID.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdExit.Enabled = True
                CmdDelete.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                s = txtSerialNo
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
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
        txtUName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdExit.Enabled = False
        CmdDelete.Enabled = False
        cmdOpen.Enabled = False
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdExit.Enabled = True
                CmdDelete.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                rsfactory.Requery

                Dim s As String
                s = txtSerialNo
                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdOpen_Click()
    frmUserSearch.Show vbModal
    cmdOpen.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "CmbHouseName") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from SUser", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If

    If Not rsfactory.EOF Then FindRecord
    
    txtSerialNo.Enabled = False

    cboEn.AddItem "1"
    cboEn.AddItem "2"
End Sub


Private Sub allenable()
    txtUName.Enabled = True
    txtPassword.Enabled = True
    txtConPas.Enabled = True
    cboEn.Enabled = True
    
End Sub

Private Sub alldisable()
    txtUName.Enabled = False
    txtPassword.Enabled = False
    txtConPas.Enabled = False
    cboEn.Enabled = False
    End Sub

Private Sub allClear()
txtSerialNo.text = ""
     txtUName.text = ""
     txtPassword.text = ""
     txtConPas.text = ""
     cboEn.text = ""
End Sub

Private Function rcupdate() As Boolean
'    On Error GoTo ErrHandler
    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then

   cn.Execute " INSERT INTO SUser(SerialNo,UID, Password, Privilegegroup ) " & _
            " VALUES ('" & txtSerialNo.text & "','" & txtUName.text & "','" & txtPassword.text & "', " & _
            IIf(cboEn = "Admin", 1, 0) & ")"

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else

        cn.Execute " UPDATE SUser SET " & _
             " UID='" & txtUName.text & "',Password= '" & txtPassword.text & "',Privilegegroup=" & _
               IIf(cboEn = "Admin", 1, 0) & " " & _
               " WHERE SerialNo='" & txtSerialNo.text & "'"
        
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
'    Exit Sub
    Exit Function

End Function
Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtSerialNo = rsfactory("SerialNo")
        txtUName = rsfactory("UID")
        txtPassword = rsfactory("Password")
        cboEn = rsfactory("Privilegegroup")
'        cboEn = IIf(rsfactory("Privilegegroup") = "Admin", 1, 0)
End If
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If Trim(txtUName) = "" Then
        MsgBox "Your are missing User Name", vbInformation
        txtUName.SetFocus
        IsValidRecord = False
        Exit Function
    End If

    If Trim(txtPassword) = "" Then
        MsgBox "Your are missing Password No.", vbInformation
        txtPassword.SetFocus
        IsValidRecord = False
        Exit Function
    End If
    
    If Trim(txtConPas) = "" Then
        MsgBox "Your are missing confirm Password No.", vbInformation
        txtConPas.SetFocus
        IsValidRecord = False
        Exit Function
        End If

   If Trim(cboEn) = "" Then
        MsgBox "Your are missing Pricilege Group No.", vbInformation
        cboEn.SetFocus
        IsValidRecord = False
        Exit Function
     End If

 If Trim(txtPassword) <> Trim(txtConPas) Then
        MsgBox "Mis Match password No.", vbInformation
        cboEn.SetFocus
        IsValidRecord = False
        Exit Function
     End If

If cmdNew.Caption = "&Save" Then
'Or cmdEdit.Caption = "&Update" Then
        If rsfactory.RecordCount > 0 Then
        If rsfactory.State <> 0 Then rsfactory.Close
            rsfactory.Open "select * from SUser where upper(UID)='" & Strings.UCase(Strings.Trim(parseQuotes(txtUName))) & "'", cn

             If Not rsfactory.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtUName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
        End If
    End Function


Public Sub PopulateItem(StrID As String)
    rsfactory.MoveFirst
    rsfactory.Find "SerialNo=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

