VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmDeliveryRequision 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Delivery Requisation Form"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "frmMenuGroup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   4
      Top             =   480
      Width           =   9615
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   5205
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   9345
         _cx             =   16484
         _cy             =   9181
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   12629161
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   65280
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   12629161
         BackColorAlternate=   14737632
         GridColor       =   12629161
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   3
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMenuGroup.frx":000C
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
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
            TabIndex        =   6
            Top             =   -840
            Width           =   8175
         End
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3000
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   6360
      Picture         =   "frmMenuGroup.frx":00C1
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   7440
      Picture         =   "frmMenuGroup.frx":098B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   750
      Left            =   8535
      Picture         =   "frmMenuGroup.frx":1255
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1100
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Enter Requisation No"
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
      Left            =   240
      TabIndex        =   8
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label3 
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
      TabIndex        =   7
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmDeliveryRequision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private rsTemp                      As ADODB.Recordset
 Private rsExport                    As ADODB.Recordset
 Private rsfactory                   As New ADODB.Recordset
 Private rsItemMaster                 As ADODB.Recordset
 Private rsItemDetail                 As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  If rsTemp.State <> 0 Then rsTemp.Close
     rsTemp.Open "SELECT SerialNo,ReqDate,DeptName,CPost " & _
                 " FROM SReqMaster WHERE SerialNo LIKE '" & RTrim(txtSearch.text) & "%'", cn, adOpenStatic, adLockReadOnly
                 fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ReqDate") & vbTab & rsTemp("DeptName") & vbTab & rsTemp("CPost")
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub

Private Sub cmdOk_Click()

'-------------------ADDD----
    Dim i As Integer
    Dim strPaymentDetail As String
    Dim strPaymentMaster As String
    Set rsItemDetail = New ADODB.Recordset
    Set rsItemMaster = New ADODB.Recordset

    If rsItemMaster.State <> 0 Then rsItemMaster.Close
    strPaymentMaster = "SELECT SerialNo, ReqDate, DeptName " & _
               "From dbo.SReqMaster where SerialNo='" & parseQuotes(Me.txtSearch) & "'"
    
    
    rsItemMaster.CursorLocation = adUseClient
    rsItemMaster.Open strPaymentMaster, cn, adOpenStatic, adLockReadOnly

   
   If rsItemMaster.RecordCount = 0 Then
        MsgBox "Please Select a requisation number From the List."
        Exit Sub
    End If

   If rsItemMaster.RecordCount > 0 Then
    frmDelivery.txtReqNo = rsItemMaster!SerialNo
    frmDelivery.ReqDate = rsItemMaster!ReqDate
    frmDelivery.cmbDepartment = rsItemMaster!DeptName
    End If

    

    
    frmDelivery.fgStock.Rows = 1
    strPaymentDetail = "SELECT  SerialNo, ReqDate, DeptName,Catagory ,SubCatagory,ItemName,Qty, " & _
                "Rate,Amount,Rol,Posted,Remarks,CPosted,Unit FROM SReqDetails " & _
                "WHERE SerialNo='" & parseQuotes(Me.txtSearch) & "' order by SerialNo "
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        frmDelivery.fgStock.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
'            frmDelivery.fgStock.TextMatrix(i, 1) = rsItemDetail("SerialNo")
'            frmDelivery.fgStock.TextMatrix(i, 2) = rsItemDetail("ReqDate")
'            frmDelivery.fgStock.TextMatrix(i, 3) = rsItemDetail("DeptName")
            frmDelivery.fgStock.TextMatrix(i, 6) = rsItemDetail("Catagory")
            frmDelivery.fgStock.TextMatrix(i, 7) = rsItemDetail("SubCatagory")
            frmDelivery.fgStock.TextMatrix(i, 8) = rsItemDetail("ItemName")
            frmDelivery.fgStock.TextMatrix(i, 10) = rsItemDetail("Qty")
            frmDelivery.fgStock.TextMatrix(i, 11) = rsItemDetail("Rate")
'            frmDelivery.fgStock.TextMatrix(i, 9) = rsItemDetail("Amount")
            frmDelivery.fgStock.TextMatrix(i, 12) = rsItemDetail("Rol")
            frmDelivery.fgStock.TextMatrix(i, 13) = rsItemDetail("Posted")
            frmDelivery.fgStock.TextMatrix(i, 14) = rsItemDetail("Remarks")
            frmDelivery.fgStock.TextMatrix(i, 15) = rsItemDetail("CPosted")
            frmDelivery.fgStock.TextMatrix(i, 16) = rsItemDetail("Unit")
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
'-------------------Add END'

'    If fgExport.RowSel < 0 Then
'        MsgBox "Please Select a Customer Name From the List."
'        Exit Sub
'    End If
'
'     Call PopulateCompanySearch
   
Unload Me
Set frmSupplierSearch = Nothing
End Sub

Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
     rsTemp.Open "SELECT SerialNo,ReqDate,DeptName,CPost FROM SReqMaster where CPost='Posted'", cn, adOpenStatic, adLockReadOnly
'                 "FROM SMSStockMaster", cn, adOpenStatic, adLockReadOnly
        
            fgExport.Rows = 1
    
    While Not rsTemp.EOF
         fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("ReqDate") & vbTab & rsTemp("DeptName") & vbTab & rsTemp("CPost")
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub

'    Private Sub PopulateCompanySearch()
'        If fgExport.Row > 0 Then
'
'             frmDelivery.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
'        End If
'    End Sub

Private Sub txtSearch_Change()
 cmdFind_Click
End Sub








''SELECT     distinct SReqMaster.SerialNo, SReqMaster.ReqDate, SReqMaster.DeptName, SReqMaster.CPost,
''           SMSDeliveryMaster.Cpost
''           FROM SReqMaster,SMSDeliveryMaster where SReqMaster.SerialNo <> SMSDeliveryMaster.ReqNo
''           and SMSDeliveryMaster.ReqNo<>0 and SReqMaster.Cpost='Posted'
''







