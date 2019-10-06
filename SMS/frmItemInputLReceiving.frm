VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmItemInputLReceiving 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Search Item Input for Receiving"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frmItemInputLReceiving.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   4575
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   5535
      Begin VB.ComboBox cmbPCatagory 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2040
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
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1560
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
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Text            =   " "
         Top             =   3720
         Width           =   5175
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox Text4 
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
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.ComboBox cmbItemName 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   2040
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
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H000000FF&
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3960
         MaskColor       =   &H000000FF&
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtExpDate 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   3120
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   56754179
         CurrentDate     =   39784
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
         TabIndex        =   22
         Top             =   3480
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
         TabIndex        =   21
         Top             =   2160
         Width           =   1695
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
         TabIndex        =   20
         Top             =   1440
         Width           =   1695
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   120
         Width           =   1695
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
         TabIndex        =   17
         Top             =   1560
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
         Left            =   3480
         TabIndex        =   16
         Top             =   2040
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
         TabIndex        =   15
         Top             =   2400
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
         Left            =   3120
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblExpireDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Expire Date"
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
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label chkWarrenty 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Warrenty"
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
         TabIndex        =   12
         Top             =   840
         Width           =   975
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
      Left            =   2520
      Picture         =   "frmItemInputLReceiving.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   735
      Left            =   1440
      Picture         =   "frmItemInputLReceiving.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin MSAdodcLib.Adodc dcItemName 
      Height          =   360
      Left            =   3600
      Top             =   5400
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dcItemName"
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
   Begin MSAdodcLib.Adodc dcCatagory 
      Height          =   360
      Left            =   3600
      Top             =   5040
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   635
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dcCatagory"
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
      TabIndex        =   23
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmItemInputLReceiving"
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

' If KeyCode = 13 Then
'        SendKeys Chr(9)
'    End If
'Call AllClear
End Sub

Private Sub cmbPCatagory_DropDown()
Call allClear

' If KeyCode = 13 Then
'        SendKeys Chr(9)
'    End If
End Sub

Private Sub cmdAdd_Click()

On Error GoTo ErrHandler

If IsValidRecord Then
 Dim iRows As Integer
 Dim i As Integer
 If txtQty = "" Then
   txtQty = 0
 End If
If txtRate = "" Then
   txtRate = 0
 End If
 
 frmStock.fgStock.AddItem "" & vbTab & vbTab & vbTab & vbTab & frmItemInputLReceiving.cmbPCatagory.text & vbTab & frmItemInputLReceiving.txtSubcode.text & _
                    vbTab & frmItemInputLReceiving.cmbItemName.text & vbTab & frmItemInputLReceiving.txtQty & vbTab & frmItemInputLReceiving.txtRate & _
                    vbTab & frmItemInputLReceiving.txtQty * frmItemInputLReceiving.txtRate & vbTab & frmItemInputLReceiving.Text4 & vbTab & frmItemInputLReceiving.dtExpDate & _
                    vbTab & vbTab & frmItemInputLReceiving.Chk1.Value & vbTab & frmItemInputLReceiving.txtRemarks.text & vbTab & vbTab & frmItemInputLReceiving.txtUnit.text


cmbItemName.RemoveItem (cmbItemName.ListIndex)
End If
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
        MsgBox "Your are missing Item Quentity Information.", vbInformation
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
    End If
  End Function


Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Call Connect
     Call ProductCatagory
     Call Itemname
     dtExpDate.Value = Null
     

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
Text4 = rsTemp1!Rol
txtUnit = rsTemp1!Unit
End Sub




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



Private Sub txtQty_GotFocus()
If IsValidRecord1 Then
Call Subcode
Call Balance
End If
End Sub

'Private Sub txtQty_Change()
''Call Subcode
''Call Balance
'End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

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
'
'rsBalance.Open ("SELECT isnull(SUM(SStockDetails.Quentity),0)as Balance FROM SStockDetails WHERE SStockDetails.ConPpsted='Posted' " & _
'               "and SStockDetails.ItemName ='" & parseQuotes(cmbItemName.text) & "'"), cn, adOpenStatic


rsBalance.Open ("select (isnull(sum(SStockDetails.Quentity),0))-(select isnull(sum(SSalesDetails.DQty),0) " & _
                "from SSalesDetails where SSalesDetails.CPosted='Posted' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
                "as Balance from SStockDetails  where SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "'  and  SStockDetails.ConPpsted='Posted'"), cn, adOpenStatic

 


If rsBalance.RecordCount > 0 Then
'      rsTemp1.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
Text16 = rsBalance!Balance

If rsBalance!Balance < 0 Then
Text16 = 0
End If


End Sub


Private Sub allClear()
txtSubcode = ""
Text16 = ""
txtUnit = ""
Text4 = ""
txtQty = ""
txtRate = ""
txtRemarks = ""
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

End Sub


