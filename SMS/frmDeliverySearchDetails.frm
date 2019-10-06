VERSION 5.00
Begin VB.Form frmDeliverySearchDetails 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Delivery Search Information"
   ClientHeight    =   7140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form4"
   ScaleHeight     =   7140
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   750
      Left            =   7575
      Picture         =   "frmDeliverySearchDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   6480
      Picture         =   "frmDeliverySearchDetails.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   5400
      Picture         =   "frmDeliverySearchDetails.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   1100
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3480
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Customer Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9615
      Begin VB.PictureBox fgExport 
         BackColor       =   &H00C0B4A9&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   5085
         Left            =   120
         ScaleHeight     =   5025
         ScaleWidth      =   9285
         TabIndex        =   1
         Top             =   360
         Width           =   9345
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
            TabIndex        =   2
            Top             =   -840
            Width           =   8175
         End
      End
   End
   Begin VB.Label Label3 
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
      TabIndex        =   8
      Top             =   0
      Width           =   9735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Enter Supplier  Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   6360
      Width           =   2775
   End
End
Attribute VB_Name = "frmDeliverySearchDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsExport                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  If rsTemp.State <> 0 Then rsTemp.Close
     rsTemp.Open "SELECT SerialNo,ReqNo,ReqDate,DeliveryDate,DepName,Cpost " & _
                 " FROM SMSDeliveryMaster WHERE SupplierName LIKE '" & RTrim(txtSearch.Text) & "%'", cn, adOpenStatic, adLockReadOnly
                 fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ReqNo") & vbTab & rsTemp("ReqDate") & vbTab & rsTemp("ReceivingDate") & vbTab & rsTemp("DepName") & vbTab & rsTemp("Cpost")
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub

Private Sub cmdOK_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Customer Name From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmCustomerSearch = Nothing
End Sub

Private Sub fgExport_DblClick()
    cmdOK_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
     rsTemp.Open "SELECT * FROM SMSDeliveryMaster", cn, adOpenStatic, adLockReadOnly
        
            fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ReqNo") & vbTab & rsTemp("ReqDate") & vbTab & rsTemp("DeliveryDate") & vbTab & rsTemp("DepName") & vbTab & rsTemp("Cpost")
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub

    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmDelivery.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub txtSearch_Change()
 cmdFind_Click
End Sub














