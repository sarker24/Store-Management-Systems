VERSION 5.00
Begin VB.Form frmItemReqReceiving 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Item Requisation"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5655
   Icon            =   "frmPaymentMode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   5535
      Begin VB.TextBox txtTAmount 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   3000
         TabIndex        =   22
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   2280
         Width           =   2265
      End
      Begin VB.ComboBox cmbPCatagory 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2265
      End
      Begin VB.ComboBox cmbItemName 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2265
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   360
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtUnit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Text            =   " "
         Top             =   3000
         Width           =   5175
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   1560
         Width           =   2265
      End
      Begin VB.TextBox txtRol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   2280
         Width           =   2265
      End
      Begin VB.TextBox txtSubcode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   480
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label lblTAmount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   2040
         Width           =   2265
      End
      Begin VB.Label lblRemarks 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label lbl 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   2265
      End
      Begin VB.Label lblQuantity 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2265
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblPCatagory 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Product Catagory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2265
      End
      Begin VB.Label lblBalance 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Banlance = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   720
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
         Height          =   375
         Left            =   3600
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblSubcode 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Subcode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      Picture         =   "frmPaymentMode.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   1320
      Picture         =   "frmPaymentMode.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1100
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
      TabIndex        =   19
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmItemReqReceiving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsTemp1                     As ADODB.Recordset
Private rsTemp2                     As ADODB.Recordset
Private rsBalance                   As ADODB.Recordset
Private bRecordExists              As Boolean
'Private rsExport                    As ADODB.Recordset
'Private rsfactory                   As New ADODB.Recordset
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdFind_Click()
''frmLedgerParty.Show vbModal
'
'        If rsTemp.State <> 0 Then rsTemp.Close
'
'
''            rsTemp.CursorLocation = adUseClient
'     rsTemp.Open "SELECT SerialNo,MenuGroup,MenuCatagory " & _
'                 "FROM tbItemlGroupSetup WHERE tbItemlGroupSetup.MenuGroup LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
'
''   End If
'
'         fgExport.Rows = 1
'
'    While Not rsTemp.EOF
'        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("MenuGroup") & _
'         vbTab & rsTemp("MenuCatagory")

'Private Sub cmbItemName_Click()
'Call Itemname
'End Sub
'Private Sub cmbItemName_Change()
'Call Itemname
'End Sub
Private Sub cmbItemName_DropDown()
Call Itemname
'Call AllClear
End Sub





Private Sub cmbPCatagory_DropDown()
Call allClear
End Sub

Private Sub cmdAdd_Click()

On Error GoTo ErrHandler

If IsValidRecord Then
'  If IsValidRecord11 Then
  
 Dim iRows As Integer
 Dim i As Integer
 If txtQty = "" Then
   txtQty = 0
 End If
If txtRate = "" Then
   txtRate = 0
 End If
If Text16 - txtQty < txtRol Then
        MsgBox "Balence is going lower level.", vbInformation
End If

frmRequisition.fgStock.AddItem "" & vbTab & vbTab & vbTab & vbTab & frmItemReqReceiving.cmbPCatagory.text & vbTab & frmItemReqReceiving.txtSubcode.text & _
                                   vbTab & frmItemReqReceiving.cmbItemName.text & vbTab & frmItemReqReceiving.txtQty & vbTab & frmItemReqReceiving.txtRate & _
                                   vbTab & frmItemReqReceiving.txtQty * frmItemReqReceiving.txtRate & vbTab & frmItemReqReceiving.txtRol & vbTab & vbTab & frmItemReqReceiving.txtRemarks.text & vbTab & vbTab & frmItemReqReceiving.txtUnit.text



'                    vbTab & frmItemReqReceiving.txtQty * frmItemReqReceiving.txtRate & vbTab & frmItemReqReceiving.dtExpDate & _
'                    vbTab & vbTab & frmItemReqReceiving.Chk1.Value & vbTab & frmItemReqReceiving.txtRemarks.text & vbTab & vbTab & frmItemReqReceiving.txtUnit.text



'
'End If

cmbItemName.RemoveItem (cmbItemName.ListIndex)
End If
'  End If
'cmbItemName.Items.Remove (cmbItemName.SelectedItem)

cmbItemName.Refresh


ErrHandler:

    Select Case Err.Number
'        Case -2147217900
        Case 13
            MsgBox "Please select numeric number in QTY/RATE field", vbInformation, "Confirmation"

   End Select
   
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    If Trim(cmbPCatagory) = "" Then
        MsgBox "Your are missing Catagory Information.", vbInformation
        cmbPCatagory.SetFocus
        IsValidRecord = False
        Exit Function
        
 ElseIf Trim(cmbItemName) = "" Then
        MsgBox "Your are missing Item Information.", vbInformation
        cmbItemName.SetFocus
        IsValidRecord = False
        Exit Function
        
  ElseIf Trim(txtQty) = "" Then
        MsgBox "Your are missing Item Qty Information.", vbInformation
        txtQty.SetFocus
        IsValidRecord = False
        Exit Function
        
   ElseIf Trim(txtRate) = "" Then
        MsgBox "Your are missing Item Rate Information.", vbInformation
        txtRate.SetFocus
        IsValidRecord = False
        Exit Function
   End If
  End Function


Private Function IsValidRecord1() As Boolean
    IsValidRecord1 = True
    
    If Trim(cmbPCatagory) = "" Then
        MsgBox "Your are missing Catagory Information.", vbInformation
        cmbPCatagory.SetFocus
        IsValidRecord1 = False
        Exit Function
        
 ElseIf Trim(cmbItemName) = "" Then
        MsgBox "Your are missing Item Information.", vbInformation
        cmbItemName.SetFocus
        IsValidRecord1 = False
        Exit Function
        
'
'ElseIf txtQty - txtRol < 0 Then
'        MsgBox "There is no Balance Qty.", vbInformation
'        cmbPCatagory.SetFocus
'        IsValidRecord1 = False
'        Exit Function
        
    End If
  End Function

Private Function IsValidRecord11() As Boolean
    IsValidRecord11 = True

    If Text16 - txtQty < txtRol Then
        MsgBox "Balence is going lower level.", vbInformation
        cmbPCatagory.SetFocus
        IsValidRecord11 = False
        Exit Function
 End If
  End Function
        
Private Sub Command2_Click()
Unload Me
End Sub

'Private Sub cmdOK_Click()
'    If fgExport.RowSel < 0 Then
'        MsgBox "Please Select a Menu Group From the List."
'        Exit Sub
'    End If
'
'     Call PopulateCompanySearch
'
'Unload Me
'Set frmItemGroupSearch = Nothing
'End Sub
'
'
'
'Private Sub fgExport_DblClick()
'    cmdOK_Click
'End Sub
'
Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Call Connect
     Call ProductCatagory
     Call Itemname
'     dtExpDate.Value = Null
     
'     Set rsTemp = New ADODB.Recordset
'     rsTemp.CursorLocation = adUseClient
'
'        If rsTemp.State <> 0 Then rsTemp.Close
'
'
'
'     rsTemp.Open "SELECT SerialNo,MenuGroup,MenuCatagory FROM tbItemlGroupSetup", cn, adOpenStatic, adLockReadOnly
'
'
'
''         fgExport.Rows = 1
'
'    While Not rsTemp.EOF
'        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("MenuGroup") & _
'         vbTab & rsTemp("MenuCatagory")
'
'        rsTemp.MoveNext
'    Wend
'     GridCount fgExport
End Sub




Private Sub Subcode()
Dim rsTemp1 As New ADODB.Recordset
If rsTemp1.State <> 0 Then rsTemp1.Close
rsTemp1.Open ("SELECT ItemName, CatagoryName,SubCatagoryCode,Rol,Unit FROM SSubCatagoryDetail where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic
'SELECT ItemName, CatagoryName,SubCatagoryCode FROM SSubCatagoryDetail where CatagoryName='" & parseQuotes(cmbPCatagory.text) & "'ORDER BY ItemName ASC
If rsTemp1.RecordCount > 0 Then
'      rsTemp1.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
txtSubcode.text = rsTemp1!SubCatagoryCode
txtRol = rsTemp1!Rol
txtUnit = rsTemp1!Unit
End Sub


'Private Sub Balance()
'Dim rsBalance As New ADODB.Recordset
'If rsBalance.State <> 0 Then rsTemp1.Close
'rsBalance.Open ("SELECT ItemName, CatagoryName,SubCatagoryCode,Rol FROM SSubCatagoryDetail where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic
''SELECT ItemName, CatagoryName,SubCatagoryCode FROM SSubCatagoryDetail where CatagoryName='" & parseQuotes(cmbPCatagory.text) & "'ORDER BY ItemName ASC
'If rsTemp1.RecordCount > 0 Then
''      rsTemp1.MoveFirst
'        bRecordExists = True
'    Else
'        bRecordExists = False
'    End If
'txtSubcode.text = rsTemp1!SubCatagoryCode
'Text4 = rsTemp1!Rol
'
'End Sub


Private Sub ProductCatagory()
     Dim rsTemp2 As New ADODB.Recordset
     
     
     rsTemp2.Open ("SELECT DISTINCT SCName as CatagoryName FROM SCatagory ORDER BY CatagoryName ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbPCatagory.AddItem rsTemp2("CatagoryName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
     
'     cmbItemName.RemoveItem (1)
     
     
'     dcCatagory.CursorLocation = adUseClient
'     dcCatagory.ConnectionString = cn.ConnectionString
'     dcCatagory.LockType = adLockReadOnly
'     dcCatagory.RecordSource = "SELECT DISTINCT SCName as CatagoryName FROM SCatagory ORDER BY CatagoryName ASC"
'     cmbPCatagory.DataMode = ssDataModeBound
'     Set cmbPCatagory.DataSource = dcCatagory
'     cmbPCatagory.DataSourceList = dcCatagory
'     cmbPCatagory.DataFieldList = "CatagoryName"
'     cmbPCatagory.DataField = "CatagoryName"
'     cmbPCatagory.ColumnHeaders = True
'     cmbPCatagory.BackColorEven = &HFFC0C0
'     cmbPCatagory.ForeColorEven = &H80000008
'     cmbPCatagory.Columns(0).Width = TextHeight("W") * 10
End Sub
Private Sub Itemname()
cmbItemName.Clear
Dim rsTemp As New ADODB.Recordset
'     dcItemName.CursorLocation = adUseClient
'     dcItemName.ConnectionString = cn.ConnectionString
'     dcItemName.LockType = adLockReadOnly
'     dcItemName.RecordSource = "SELECT DISTINCT ItemName,SubCatagoryCode CatagoryName FROM SSubCatagoryDetail where CatagoryName='" & parseQuotes(cmbPCatagory) & "'ORDER BY ItemName ASC"
''     dcItemName.RecordSource = "SELECT DISTINCT SubCatagoryCode,ItemName FROM SSubCatagoryDetail"
'     cmbItemName.DataMode = ssDataModeBound
'     Set cmbItemName.DataSource = dcItemName
'     cmbItemName.DataSourceList = dcItemName
'     cmbItemName.DataFieldList = "ItemName"
'     cmbItemName.DataField = "ItemName"
'     cmbItemName.Columns(2).Visible = False
'     cmbItemName.ColumnHeaders = True
'     cmbItemName.BackColorEven = &HFFC0C0
'     cmbItemName.ForeColorEven = &H80000008


rsTemp.Open ("SELECT ItemName, CatagoryName FROM SSubCatagoryDetail where CatagoryName='" & parseQuotes(cmbPCatagory.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic
    While Not rsTemp.EOF
        cmbItemName.AddItem rsTemp("ItemName")
        rsTemp.MoveNext
    Wend
    rsTemp.Close
     
'     cmbItemName.RemoveItem (1)

End Sub



Private Sub txtQty_Change()
'Call Subcode
'Call Balance
End Sub
Private Sub txtQty_Click()
If IsValidRecord1 Then
Call Subcode
Call Balance
End If
End Sub

Private Sub Balance()
Dim rsBalance As New ADODB.Recordset
If rsBalance.State <> 0 Then rsBalance.Close
'rsBalance.Open ("SELECT ItemName, CatagoryName,SubCatagoryCode,Rol FROM SSubCatagoryDetail where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic
'SELECT ItemName, CatagoryName,SubCatagoryCode FROM SSubCatagoryDetail where CatagoryName='" & parseQuotes(cmbPCatagory.text) & "'ORDER BY ItemName ASC



'rsBalance.Open ("SELECT ItemName, CatagoryName,SubCatagoryCode,Rol FROM SSubCatagoryDetail where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic

'rsBalance.Open ("SELECT isnull(SUM(SStockDetails.Qty),0)as Balance FROM SStockDetails WHERE SStockDetails.CPost='Posted' " & _
'               "and SStockDetails.ItemName ='" & parseQuotes(cmbItemName.text) & "'"), cn, adOpenStatic


rsBalance.Open ("select (isnull(sum(SStockDetails.Qty),0))-(select isnull(sum(SSalesDetails.Qty),0) " & _
                "from SSalesDetails where SSalesDetails.CPosted='Posted' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
                "as Balance from SStockDetails  where SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "'  and  SStockDetails.CPost='Posted'"), cn, adOpenStatic



If rsBalance.RecordCount > 0 Then
'      rsTemp1.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
Text16.text = rsBalance!Balance


End Sub


Private Sub allClear()
txtSubcode = ""
Text16 = ""
txtUnit = ""
txtRol = ""
txtQty = ""
txtRate = ""
'dtExpDate.Value = ""
txtRemarks = ""
'Chk1.Value = 0

End Sub

'Private Sub txtRate_Click()
'If IsValidRecord11 Then
''Call Subcode
''Call Balance
'End If
'End Sub


Private Sub txtRate_click()
If IsValidRecord2 Then
   If Text16 - txtQty < 0 Then
        MsgBox "Req Qty must lower than Balance.", vbInformation
        cmbPCatagory.SetFocus
        


'    If Text16 - txtQty < txtRol Then
'        MsgBox "Balence is going lower level.", vbInformation
'        cmbPCatagory.SetFocus
'
   
        
        
   End If
 End If
 End Sub
 Private Function IsValidRecord2() As Boolean
    IsValidRecord2 = True
    
    If Trim(cmbPCatagory) = "" Then
        MsgBox "Your are missing Catagory Information.", vbInformation
        cmbPCatagory.SetFocus
        IsValidRecord2 = False
        Exit Function
        
 ElseIf Trim(cmbItemName) = "" Then
        MsgBox "Your are missing Item Information.", vbInformation
        cmbItemName.SetFocus
        IsValidRecord2 = False
        Exit Function
        
  ElseIf Trim(txtQty) = "" Then
        MsgBox "Your are missing Item Qty Information.", vbInformation
        txtQty.SetFocus
        IsValidRecord2 = False
        Exit Function
        
'   ElseIf Trim(txtRate) = "" Then
'        MsgBox "Your are missing Item Rate Information.", vbInformation
'        txtRate.SetFocus
'        IsValidRecord = False
'        Exit Function
   End If
  End Function
 
