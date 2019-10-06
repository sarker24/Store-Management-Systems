VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmItemInputReceiving 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Item Input for Receiving"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   Icon            =   "frmItemGroupSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   1320
      Picture         =   "frmItemGroupSearch.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
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
      Height          =   735
      Left            =   2400
      Picture         =   "frmItemGroupSearch.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   4575
      Left            =   0
      TabIndex        =   8
      Top             =   360
      Width           =   5535
      Begin VB.TextBox txtLPRate 
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2160
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
         Width           =   2265
      End
      Begin VB.TextBox txtCatagory 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
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
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3360
         Width           =   2265
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
         Top             =   2760
         Width           =   2265
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
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1080
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Chk1 
         BackColor       =   &H000000FF&
         Caption         =   "Check1"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4080
         MaskColor       =   &H000000FF&
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
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
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   2160
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
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtQty 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Please Enter Only Numeric Value"
         Top             =   1560
         Width           =   2265
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Text            =   " "
         Top             =   4080
         Width           =   5175
      End
      Begin MSComCtl2.DTPicker dtExpDate 
         Height          =   315
         Left            =   3000
         TabIndex        =   7
         Top             =   3360
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   65208323
         CurrentDate     =   39784
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
         TabIndex        =   30
         Top             =   120
         Width           =   2265
      End
      Begin VB.Label lblLPRate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0B4A9&
         Caption         =   "Last P. Rate"
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
         TabIndex        =   29
         Top             =   2160
         Width           =   1095
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
         TabIndex        =   27
         Top             =   2520
         Width           =   2265
      End
      Begin VB.Label chkWarrenty 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   24
         Top             =   360
         Width           =   1095
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
         Left            =   3000
         TabIndex        =   22
         Top             =   3120
         Width           =   2265
      End
      Begin VB.Label lblSubcode 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   21
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
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
         Left            =   2880
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0B4A9&
         Caption         =   "Banlance"
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
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
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
         Top             =   3120
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1920
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
         TabIndex        =   10
         Top             =   3840
         Width           =   2265
      End
   End
   Begin MSAdodcLib.Adodc dcItemName 
      Height          =   360
      Left            =   3480
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
      Left            =   3480
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
      Left            =   -480
      TabIndex        =   19
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmItemInputReceiving"
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

Private Sub cmbItemName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbItemName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbItemName_LostFocus()
Call Others
Call Balance
'Call Itemname
Call LPRate
End Sub

Private Sub Others()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT SerialNo,Catagory,ItemName,ROL FROM SubCatagory where ItemName='" & parseQuotes(cmbItemName.text) & "'"), cn, adOpenStatic

'txtSRate = rsTemp!SRate
'txtMID = rsTemp!SerialNo
txtCatagory = rsTemp!Catagory
    
    rsTemp.Close
End Sub

Private Sub LPRate()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset
     rsTemp.Open ("SELECT TOP 1 id,ItemName, Rate FROM SStockDetails where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ID desc"), cn, adOpenStatic
    If rsTemp.RecordCount > 0 Then
    txtLPRate = rsTemp!Rate
 End If
    rsTemp.Close
End Sub


Private Sub cmbStoreName_DropDown()
Call allClear
End Sub

Private Sub cmdAdd_Click()

On Error GoTo ErrHandler

If IsValidRecord Then
 Dim iRows As Integer
 Dim i As Integer
 
 If cmbItemName = "" Then
 MsgBox "Please Input Item Name.", vbInformation
        cmbItemName.SetFocus
    End If
 
 If txtQty = "" Then
 MsgBox "Please Input Item Quantity.", vbInformation
        txtQty.SetFocus
   txtQty = 0
 End If
If txtRate = "" Then
   txtRate = 0
 End If
 
 frmStock.fgStock.AddItem "" & vbTab & vbTab & vbTab & vbTab & frmItemInputReceiving.cmbStoreName.text & vbTab & frmItemInputReceiving.txtCatagory.text & _
                    vbTab & frmItemInputReceiving.cmbItemName.text & vbTab & frmItemInputReceiving.txtQty & vbTab & frmItemInputReceiving.txtRate & _
                    vbTab & frmItemInputReceiving.txtQty * frmItemInputReceiving.txtRate & vbTab & frmItemInputReceiving.txtROL & vbTab & frmItemInputReceiving.dtExpDate & _
                    vbTab & vbTab & frmItemInputReceiving.Chk1.Value & vbTab & frmItemInputReceiving.txtRemarks.text & vbTab & vbTab & frmItemInputReceiving.txtUnit.text


cmbItemName.RemoveItem (cmbItemName.ListIndex)
End If
'cmbItemName.Items.Remove (cmbItemName.SelectedItem)

cmbItemName.Refresh


ErrHandler:

    Select Case Err.Number
'        Case -2147217900
        Case 13
            MsgBox "Please Inpur Numeric Number.", vbInformation, "Confirmation"

   End Select
   
   allClear
   cmbItemName.SetFocus
   
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    If Trim(cmbItemName) = "" Then
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
    
       
 If Trim(cmbItemName) = "" Then
        MsgBox "Your are missing Item Information.", vbInformation
        cmbItemName.SetFocus
        IsValidRecord1 = False
        Exit Function
    End If
  End Function


Private Sub CmdExit_Click()
Unload Me
End Sub


Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Call Connect
     Call StoreName
     Call Itemname
     dtExpDate.Value = Date
     

End Sub

Private Sub Subcode()
Dim rsTemp1 As New ADODB.Recordset
If rsTemp1.State <> 0 Then rsTemp1.Close
rsTemp1.Open ("SELECT ItemName, Catagory,Rol,Unit FROM SubCatagory where ItemName='" & parseQuotes(cmbItemName.text) & "'ORDER BY ItemName ASC"), cn, adOpenStatic
'SELECT ItemName, CatagoryName,SubCatagoryCode FROM SSubCatagoryDetail where CatagoryName='" & parseQuotes(cmbPCatagory.text) & "'ORDER BY ItemName ASC
If rsTemp1.RecordCount > 0 Then
'      rsTemp1.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
'txtSubcode.text = rsTemp1!SubCatagoryCode
txtROL = rsTemp1!Rol
txtUnit = rsTemp1!Unit
txtCatagory = rsTemp1!Catagory
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
    
    While Not rsTemp2.EOF
        cmbItemName.AddItem rsTemp2("ItemName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
     
End Sub

Private Sub txtQty_GotFocus()
If IsValidRecord1 Then
Call Subcode
Call Balance
End If
End Sub

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

rsBalance.Open ("select (isnull(sum(SStockDetails.Qty),0))-(select isnull(sum(SSalesDetails.Qty),0) " & _
                "from SSalesDetails where SSalesDetails.CPost='Posted'and SSalesDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SSalesDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' ) " & _
                "as Balance from SStockDetails  where SStockDetails.StoreName='" & parseQuotes(cmbStoreName.text) & "' and SStockDetails.ItemName='" & parseQuotes(cmbItemName.text) & "' and  SStockDetails.CPost='Posted'"), cn, adOpenStatic

If rsBalance.RecordCount > 0 Then
'      rsTemp1.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
txtBalance = rsBalance!Balance

If rsBalance!Balance < 0 Then
txtBalance = 0
End If


End Sub


Private Sub allClear()
txtSubcode = ""
txtBalance = ""
txtUnit = ""
txtROL = ""
txtQty = ""
txtRate = ""
txtTAmount = ""
txtLPRate = ""
txtCatagory = ""
txtRemarks = ""

End Sub

Private Sub txtRate_Change()
Call TRCalculation
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If
    
    Call TRCalculation
'    cmdAdd.SetFocus

End Sub

Private Sub txtRate_LostFocus()
Call TRCalculation
End Sub

Private Sub txtTAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
Call RCalculation
'Call TRCalculation

End Sub
Private Sub TRCalculation()
txtTAmount = CDbl(Val(txtQty)) * CDbl(Val(txtRate))
End Sub

Private Sub RCalculation()
txtRate = CDbl(Val(txtTAmount)) / CDbl(Val(txtQty))
End Sub


Private Sub txtTAmount_LostFocus()
Call RCalculation
End Sub
