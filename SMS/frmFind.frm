VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmStockSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   " Stock Master Search "
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9870
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMode 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   " "
      Top             =   6480
      Width           =   2295
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   6480
      Width           =   2655
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
      TabIndex        =   3
      Top             =   480
      Width           =   9615
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   5205
         Left            =   120
         TabIndex        =   4
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFind.frx":000C
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
            TabIndex        =   5
            Top             =   -840
            Width           =   8175
         End
      End
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   6360
      Picture         =   "frmFind.frx":0110
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
      Picture         =   "frmFind.frx":09DA
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
      Picture         =   "frmFind.frx":12A4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1100
   End
   Begin VB.Label lblIteamGroup 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Field Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Available Value"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   6240
      Width           =   2535
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
      TabIndex        =   6
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "frmStockSearch"
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
  
  If cboMode.text = "Serial Number" Then

      rsTemp.Open "SELECT TOP 50 SerialNo, RDate, SName, SBill, DName, StoreName, Posted " & _
                 "FROM SStockMaster WHERE SStockMaster.SerialNo LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
ElseIf cboMode.text = "Store Name" Then

      rsTemp.Open "SELECT TOP 50 SerialNo, RDate, SName, SBill, DName, StoreName, Posted " & _
                 "FROM SStockMaster WHERE SStockMaster.StoreName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Department Name" Then

rsTemp.Open "SELECT TOP 50 SerialNo, RDate, SName, SBill, DName, StoreName, Posted " & _
            "FROM SStockMaster WHERE SStockMaster.DName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

Else
 
 rsTemp.Open "SELECT TOP 50 SerialNo, RDate, SName, SBill, DName, StoreName, Posted " & _
             "FROM SStockMaster WHERE SStockMaster.Posted LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

End If
    
    fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("RDate") & vbTab & rsTemp("SName") & vbTab & rsTemp("SBill") & vbTab & rsTemp("DName") & vbTab & rsTemp("StoreName") & vbTab & rsTemp("Posted")
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Store Name From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmStockSearch = Nothing

'Set frmCustomerSearch = Nothing
End Sub

Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
     rsTemp.Open "SELECT * FROM SStockMaster", cn, adOpenStatic, adLockReadOnly
'                 "FROM SStockMaster", cn, adOpenStatic, adLockReadOnly
        
            fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("RDate") & vbTab & rsTemp("SName") & vbTab & rsTemp("SBill") & vbTab & rsTemp("DName") & vbTab & rsTemp("StoreName") & vbTab & rsTemp("Posted")
        rsTemp.MoveNext
    Wend
     GridCount fgExport
     
cboMode.AddItem "Serial Number"
cboMode.AddItem "Store Name"
cboMode.AddItem "Department Name"
cboMode.AddItem "Supplier Bill"
cboMode.text = "Serial Number"

End Sub

    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmStock.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub txtSearch_Change()
 cmdFind_Click
End Sub
