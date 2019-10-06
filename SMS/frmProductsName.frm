VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmReceivingProduct 
   BackColor       =   &H00C0B4A9&
   Caption         =   " Receiving  Products Information"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   Icon            =   "frmProductsName.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12450
   ScaleWidth      =   17160
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   615
      Left            =   360
      Picture         =   "frmProductsName.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   615
      Left            =   1320
      Picture         =   "frmProductsName.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      Height          =   615
      Left            =   2280
      Picture         =   "frmProductsName.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3240
      Picture         =   "frmProductsName.frx":1CA8
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Q&uit"
      Height          =   615
      Left            =   4200
      Picture         =   "frmProductsName.frx":2572
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9960
      Width           =   990
   End
   Begin VB.CommandButton chameleonButton1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      Height          =   615
      Left            =   5160
      Picture         =   "frmProductsName.frx":2E3C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9960
      Width           =   990
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Product Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   6735
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   14775
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00C0B4A9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   14400
         Picture         =   "frmProductsName.frx":3706
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Remove"
         Top             =   600
         Width           =   300
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0B4A9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   14400
         Picture         =   "frmProductsName.frx":3C90
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add"
         Top             =   240
         Width           =   300
      End
      Begin VSFlex7DAOCtl.VSFlexGrid fgProducts 
         Height          =   6135
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   14175
         _cx             =   25003
         _cy             =   10821
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   0
         ForeColorFixed  =   49152
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   12629161
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmProductsName.frx":421A
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
         DataMode        =   0
         VirtualData     =   -1  'True
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   49152
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Product Master Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   14775
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   15
         Top             =   1200
         Width           =   2535
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo2 
         Height          =   375
         Left            =   6000
         TabIndex        =   13
         Top             =   1080
         Width           =   2535
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo SSOleDBCombo1 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   360
         Width           =   2535
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   4471
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
      End
      Begin VB.Label lblSerial 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Serial No"
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
         Left            =   720
         TabIndex        =   18
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Sub Catagory"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblProductCatagory 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Product Catagory"
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
         Left            =   4320
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmReceivingProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private rs As adodb.Recordset
'
'
'Private Sub cmdAdd_Click(Index As Integer)
'Dim i As Integer, j As Integer
'Dim rs As New adodb.Recordset
''-----------------------------------------
'Select Case Index
'Case 1
'    fgProducts.AddItem ""
'    fgProducts.Col = 1
'
'    If flagSlNo = 0 Then
'    rs.Open "Select SL=isnull(max(LedgerID),0) from USLedgerDetail", cn, adOpenStatic
'    j = rs!SL + 1
'    fgProducts.TextMatrix(fgLedger.Rows - 1, 1) = j
'    flagSlNo = 1
'    Else
'       fgProducts.TextMatrix(fgLedger.Rows - 1, 1) = fgLedger.TextMatrix(fgLedger.Rows - 2, 1) + 1
'        j = j + 1
'    End If
'End Select
'End Sub
'
'Private Sub cmdDel_Click(Index As Integer)
'
'Select Case Index
'Case 0
'    If fgProducts.Rows = 1 Then Exit Sub
'    If fgProducts.Row >= 1 Then
'        fgProducts.RemoveItem fgLedger.Row
'    Else
'        MsgBox "You have to select a row to delete.", vbInformation, "General"
'    End If
'
'End Select
'End Sub
Private Sub fgProducts_Click()

End Sub

