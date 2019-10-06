VERSION 5.00
Begin VB.Form frmIteminputDelivery 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0B4A9&
   Caption         =   "Select Item For Delivery"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   Icon            =   "frmIteminput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   3480
      Picture         =   "frmIteminput.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   750
      Left            =   4560
      Picture         =   "frmIteminput.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   4215
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   5535
      Begin VB.TextBox txtFBalance 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3960
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox cmbStoreName 
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
         Width           =   2295
      End
      Begin VB.TextBox txtTAmount 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   3000
         Width           =   2265
      End
      Begin VB.TextBox txtCatagory 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2415
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
         Height          =   360
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   435
         Width           =   1215
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   2265
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1215
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
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3600
         Width           =   5175
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
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtBalance 
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Balance"
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
         Left            =   2640
         TabIndex        =   27
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblStoreName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Store Name"
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
         TabIndex        =   25
         Top             =   120
         Width           =   2265
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
         Left            =   2880
         TabIndex        =   24
         Top             =   2760
         Width           =   2265
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
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   2265
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
         Left            =   3120
         TabIndex        =   19
         Top             =   480
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
         Left            =   3480
         TabIndex        =   18
         Top             =   1560
         Width           =   615
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
         Left            =   3480
         TabIndex        =   17
         Top             =   1200
         Width           =   735
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
         Left            =   3120
         TabIndex        =   16
         Top             =   840
         Width           =   855
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
         TabIndex        =   14
         Top             =   1320
         Width           =   2265
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
         TabIndex        =   13
         Top             =   2025
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
         TabIndex        =   12
         Top             =   3360
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      TabIndex        =   21
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmIteminputDelivery"
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

Private Sub cmbItemName_DropDown()
Call Itemname
End Sub

Private Sub cmbItemName_LostFocus()
'Call Subcode
Call Balance
'Call PRate
End Sub

Private Sub cmbItemName_KeyPress(KeyAscii As Integer)

KeyAscii = AutoMatchCBBox(cmbItemName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If

End Sub

Private Sub cmbStoreName_DropDown()
Call allClear
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Call Connect
     Call StoreName
     Call Itemname
End Sub

Private Sub StoreName()
     

Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT StoreName FROM HStore ORDER BY StoreName ASC"), cn, adOpenStatic
While Not rsTemp.EOF
cmbStoreName.AddItem rsTemp("StoreName")
rsTemp.MoveNext
Wend
rsTemp.Close
    
'    Call Others
     
End Sub

Private Sub Itemname()
     cmbItemName.Clear
     Dim rsTemp2 As New ADODB.Recordset
     
     rsTemp2.Open ("SELECT StoreName,ItemName FROM SubCatagory where StoreName='" & parseQuotes(cmbStoreName.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic
'   rsTemp2.Open "exec Input_Delivery_Item", cn, adOpenStatic
    While Not rsTemp2.EOF
        cmbItemName.AddItem rsTemp2("ItemName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
     
End Sub
'Private Sub PRate()
'On Error Resume Next
'
'     Dim rsTemp As New ADODB.Recordset
'     rsTemp.Open ("SELECT TOP 1 id,ItemName, Rate FROM SStockDetails where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ID desc"), cn, adOpenStatic
'    If rsTemp.RecordCount > 0 Then
'    txtRate = rsTemp!Rate
' End If
'    rsTemp.Close
'End Sub

Private Sub cmdAdd_Click()

On Error GoTo ErrHandler

If IsValidRecord Then
 Dim iRows As Integer
 Dim i As Integer
 If txtQty = "" Then
'  MsgBox "Input Item Qty Information.", vbInformation
'   txtQty.SetFocus
   txtQty = 0
 End If
If txtRate = "" Then
   txtRate = 0
 End If
 
 If txtBalance - txtQty < txtROL Then
        MsgBox "Balance is going lower level.", vbInformation
End If

frmDelivery.fgStock.AddItem "" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & frmIteminputDelivery.cmbStoreName.text & vbTab & frmIteminputDelivery.txtCatagory.text & _
                    vbTab & frmIteminputDelivery.cmbItemName.text & vbTab & frmIteminputDelivery.txtBalance & vbTab & frmIteminputDelivery.txtQty & _
                    vbTab & frmIteminputDelivery.txtRate & vbTab & frmIteminputDelivery.txtQty * frmIteminputDelivery.txtRate & vbTab & vbTab & frmIteminputDelivery.txtROL.text & _
                    vbTab & vbTab & frmIteminputDelivery.txtRemarks.text & vbTab & vbTab & frmIteminputDelivery.txtUnit.text


cmbItemName.RemoveItem (cmbItemName.ListIndex)
End If

cmbItemName.Refresh

ErrHandler:

    Select Case Err.Number
'        Case -2147217900
        Case 13
            MsgBox "Please select numeric number in QTY/RATE field", vbInformation, "Confirmation"
   End Select
   
   Call allClear
   cmbItemName.SetFocus
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
If Trim(cmbItemName) = "" Then
        MsgBox "Select Item Name Information.", vbInformation
        cmbItemName.SetFocus
        IsValidRecord = False
        Exit Function
        
  ElseIf Trim(txtQty) = "" Then
        MsgBox "Input Item Qty Information.", vbInformation
        txtQty.SetFocus
        IsValidRecord = False
        Exit Function
        
 ElseIf Val(txtQty) > Val(txtBalance) Then
        MsgBox "Product Item Stock are not available.", vbInformation
        txtQty.SetFocus
        IsValidRecord = False
        Exit Function
   End If
  End Function

Private Function IsValidRecord1() As Boolean
    IsValidRecord1 = True
    
If Trim(cmbItemName) = "" Then
        MsgBox "Select Item Name Information.", vbInformation
        cmbItemName.SetFocus
        IsValidRecord1 = False
        Exit Function
        
ElseIf Val(txtQty) > Val(txtBalance) Then
        MsgBox "Product Item Stock are not available.", vbInformation
        txtQty.SetFocus
        IsValidRecord1 = False
        Exit Function
   End If
  End Function


Private Sub CmdExit_Click()
Unload Me
End Sub


Private Sub Calculation()
txtTAmount = Val(txtQty) * Val(txtRate)
End Sub

Private Sub Subcode()

Dim rsTemp1 As New ADODB.Recordset
If rsTemp1.State <> 0 Then rsTemp1.Close
rsTemp1.Open ("SELECT ItemName, Catagory,Rol,Unit FROM SubCatagory where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic
If rsTemp1.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
'txtSubcode.text = rsTemp1!SubCatagoryCode
txtROL = rsTemp1!Rol
txtUnit = rsTemp1!Unit
txtCatagory = rsTemp1!Catagory

End Sub

Private Sub txtQty_Change()
If IsValidRecord1 Then
Call Subcode
Call Balance
Call Calculation
TotalBalance
End If
End Sub

Private Sub txtQty_GotFocus()
If IsValidRecord1 Then
Call Subcode
Call Balance
TotalBalance
End If
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

End Sub


Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub Balance()

Dim rsBalance As New ADODB.Recordset
If rsBalance.State <> 0 Then rsBalance.Close
txtBalance.text = ""

'rsBalance.Open ("select (isnull(sum(SStockDetails.Qty),0))-(select isnull(sum(SSalesDetails.Qty),0) " & _
'               "from SSalesDetails where SSalesDetails.CPost='Posted'and SSalesDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
'               "as Balance from SStockDetails  where SStockDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' and  SStockDetails.CPost='Posted'"), cn, adOpenStatic
'


rsBalance.Open "exec Input_Delivery_Item '" & Trim(cmbStoreName.text) & "','" & parseQuotes(cmbItemName.text) & "'", cn, adOpenStatic

If rsBalance.RecordCount > 0 Then
'      rsTemp1.MoveFirst
       bRecordExists = True
   Else
       bRecordExists = False
   End If
'txtBalance.text = rsBalance!qty

If rsBalance.EOF Then

Else
txtBalance.text = rsBalance!qty
'txtBalance.text = rsBalance!Balance
txtRate.text = rsBalance!Rate
End If

'End Sub

'Dim rsBalance As New ADODB.Recordset
'If rsBalance.State <> 0 Then rsBalance.Close
'
'rsBalance.Open ("select (isnull(sum(SStockDetails.Qty),0))-(select isnull(sum(SSalesDetails.Qty),0) " & _
'               "from SSalesDetails where SSalesDetails.CPost='Posted'and SSalesDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
'               "as Balance from SStockDetails  where SStockDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' and  SStockDetails.CPost='Posted'"), cn, adOpenStatic
'
''rsBalance.Open ("select (isnull(sum(SStockDetails.Qty),0))-(select isnull(sum(SSalesDetails.Qty),0) " & _
''                "from SSalesDetails where SSalesDetails.CPost='Posted' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
''                "as Balance from SStockDetails  where SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "'  and  SStockDetails.CPost='Posted'"), cn, adOpenStatic
'
'
'If rsBalance.RecordCount > 0 Then
''      rsTemp1.MoveFirst
'       bRecordExists = True
'   Else
'       bRecordExists = False
'   End If
'txtBalance.text = rsBalance!Balance


End Sub


Private Sub TotalBalance()

        Dim rsBalance As New ADODB.Recordset
        If rsBalance.State <> 0 Then rsBalance.Close
        
        rsBalance.Open ("select (isnull(sum(SStockDetails.Qty),0))-(select isnull(sum(SSalesDetails.Qty),0) " & _
                       "from SSalesDetails where SSalesDetails.CPost='Posted'and SSalesDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
                       "as Balance from SStockDetails  where SStockDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' and  SStockDetails.CPost='Posted'"), cn, adOpenStatic
        
        'rsBalance.Open ("select (isnull(sum(SStockDetails.Qty),0))-(select isnull(sum(SSalesDetails.Qty),0) " & _
        '                "from SSalesDetails where SSalesDetails.CPost='Posted' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
        '                "as Balance from SStockDetails  where SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "'  and  SStockDetails.CPost='Posted'"), cn, adOpenStatic
        
        
        If rsBalance.RecordCount > 0 Then
        '      rsTemp1.MoveFirst
               bRecordExists = True
           Else
               bRecordExists = False
           End If
        txtFBalance.text = rsBalance!Balance
End Sub

Private Sub allClear()
txtSubcode = ""
txtBalance = ""
txtUnit = ""
txtROL = ""
txtQty = ""
txtRate = ""
txtTAmount = ""
txtRemarks = ""
txtFBalance = ""

End Sub

'Private Sub txtRate_click()
'If IsValidRecord2 Then
'    If txtBalance - txtQty < 0 Then
'           MsgBox "Delivery Qty must lower than Balance.", vbInformation
'           cmbPCatagory.SetFocus
'
'    End If
' End If
'End Sub

Private Function IsValidRecord2() As Boolean
    IsValidRecord2 = True
    
If Trim(cmbItemName) = "" Then
        MsgBox "You are missing Item Information.", vbInformation
        cmbItemName.SetFocus
        IsValidRecord2 = False
        Exit Function
        
  ElseIf Trim(txtQty) = "" Then
        MsgBox "You are missing Item Quantity Information.", vbInformation
        txtQty.SetFocus
        IsValidRecord2 = False
        Exit Function
        
   End If
  End Function


