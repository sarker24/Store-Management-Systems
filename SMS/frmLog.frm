VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   0  'None
   Caption         =   "Login To application ..."
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   DrawMode        =   12  'Nop
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLog.frx":058A
   ScaleHeight     =   2835
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTime 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   3960
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   120
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   2880
      Picture         =   "frmLog.frx":3320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton CmdEnter 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   1920
      Picture         =   "frmLog.frx":3BEA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "1991"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox txtUID 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Text            =   "DEBDAS"
      Top             =   960
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   120
      Picture         =   "frmLog.frx":42D4
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1365
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BorderColor     =   &H80000001&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   1095
      Left            =   1560
      Top             =   840
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Tag             =   "&User Name:"
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Tag             =   "&Password:"
      Top             =   1440
      Width           =   960
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub CmdEnter_Click()
    Call Connect
    Dim str As String
    Dim cm As New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set cm = New ADODB.Connection '
    Set cn = New ADODB.Connection
'   str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & SDatabaseName & ";Data Source=" & sServerName
    str = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=" & SDatabaseName & ";Server=" & sServerName
    cn.Open str
    txtDate.text = Date        'Format((cm.Execute("Select GetDate()")), "dd-MM-yyyy")
    txtTime.text = Time        'Format((cm.Execute("Select GetDate()")), "hh:mm:ss")
    str = "select UID,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & txtUID.text & "'"
'    str = "select SerialNo,Password,Privilegegroup,Upper(UID)as Name  from SUser where UID ='" & txtUID.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
    If rs!Password = Trim(CStr(txtPassword.text)) Then
'    frmScanLogOn.Show vbModal
        Call frmMain.Show
        If rs!Privilegegroup = 1 Then
            frmLogin.Hide
             frmMain.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuCommunication.Enabled = True
             frmMain.mnuExit.Enabled = True
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuReport.Enabled = True
             frmMain.mnuCalculator.Enabled = True
             
             
             frmStock.CmdDelete.Enabled = True
             frmStock.cmdUndoPost.Enabled = True
             frmRequisition.CmdDelete.Enabled = True
             frmRequisition.cmdUndoPost.Enabled = True
             frmDelivery.CmdDelete.Enabled = True
             frmDelivery.cmdUndoPost.Enabled = True
             
             
             frmSuppliername.CmdDelete.Enabled = True
             frmDepartment.CmdDelete.Enabled = True
             frmCatagorySub.CmdDelete.Enabled = True
             frmUser.Enabled = True
             frmCatagory.CmdDelete.Enabled = True
             
                      
         str = " insert into RMSLogin Values ('" & txtUID.text & "','" & Format(txtDate.text, "yyyy-mm-dd") & "','" & Format(txtTime.text, "HH:MM:SS") & "')"
         
  
  
         Else
             frmMain.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuCommunication.Enabled = False
             frmMain.mnuExit.Enabled = True
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuReport.Visible = True
             frmMain.mnuCalculator.Enabled = True
             frmMain.mnuUser.Visible = False
             
             
             frmStock.CmdDelete.Enabled = False
             frmStock.cmdUndoPost.Enabled = False
             frmRequisition.CmdDelete.Enabled = False
             frmRequisition.cmdUndoPost.Enabled = False
             frmDelivery.CmdDelete.Enabled = False
             frmDelivery.cmdUndoPost.Enabled = False
              
              
             frmSuppliername.CmdDelete.Enabled = False
             frmDepartment.CmdDelete.Enabled = False
             frmCatagorySub.CmdDelete.Enabled = False
             frmUser.Enabled = False
             frmCatagory.CmdDelete.Enabled = False
'             frmLivePurchase.cmdDelete.Enabled = False
             frmLogin.Hide
             
          
             
        End If
    Else
            MsgBox "Invalid Password. Please try again.", vbInformation, "Confarmation"
            txtPassword.text = ""
            txtPassword.SetFocus
    End If


End Sub

Private Sub CmdEnter_GotFocus()
    CmdEnter.FontBold = True
End Sub

Private Sub CmdEnter_LostFocus()
    CmdEnter.FontBold = False
End Sub

Private Sub CmdCancel_GotFocus()
    cmdCancel.FontBold = True
End Sub

Private Sub CmdCancel_LostFocus()
    cmdCancel.FontBold = False
End Sub



Private Sub txtUID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtPassword.SetFocus
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call CmdEnter_Click
    End If
End Sub

Private Sub Form_Load()
'   ModFunction.StartUpPosition Me
'   Text1.Enabled = True
End Sub

Private Sub txtUID_GotFocus()
    txtUID.BackColor = &HFFC0C0
    txtUID.SelStart = 0
    txtUID.SelLength = Len(txtUID)
End Sub

Private Sub txtUID_LostFocus()
    txtUID.BackColor = &HFFFFFF
    txtUID.text = StrConv(txtUID.text, vbProperCase)
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.BackColor = &HFFC0C0
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword.BackColor = &HFFFFFF
End Sub




