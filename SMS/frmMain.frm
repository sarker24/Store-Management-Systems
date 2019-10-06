VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000A&
   Caption         =   "Store Mangement System [RCH Mother Store]"
   ClientHeight    =   10065
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":08CA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9690
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Current User :"
            TextSave        =   "Current User :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   11853
            Text            =   "Software Developed by ""MAS IT SOLUTIONS"". Contact : +880-2-8056691, 01915682291"
            TextSave        =   "Software Developed by ""MAS IT SOLUTIONS"". Contact : +880-2-8056691, 01915682291"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MNu 
      Caption         =   "--------"
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "Setup"
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer Name"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSName 
         Caption         =   "Supplier Name"
      End
      Begin VB.Menu mnuStoreName 
         Caption         =   "Store Name"
      End
      Begin VB.Menu mnuDepartmentName 
         Caption         =   "Department Name"
      End
      Begin VB.Menu mnuProductCatagory 
         Caption         =   "Product Catagory"
      End
      Begin VB.Menu mnuProductSubCatagory 
         Caption         =   "Product Sub Catagory"
      End
      Begin VB.Menu mnuUser 
         Caption         =   "User Setup"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu mnuRequisition 
         Caption         =   "Requisition"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReceiving 
         Caption         =   "Receiving"
      End
      Begin VB.Menu mnuDelivery 
         Caption         =   "Delivery"
      End
      Begin VB.Menu mnuSPInformation 
         Caption         =   "Supplier Payment Information"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "Accounts"
      Visible         =   0   'False
      Begin VB.Menu mnuHAccounts 
         Caption         =   "Head of Accounts"
      End
      Begin VB.Menu mnuVEntry 
         Caption         =   "Voucher Entry"
      End
      Begin VB.Menu mnuFloorsheet 
         Caption         =   "Floor Sheet"
      End
      Begin VB.Menu mnuCBook 
         Caption         =   "Cash Book"
      End
      Begin VB.Menu mnuBBook 
         Caption         =   "Bank Book"
      End
      Begin VB.Menu mnuGLedger 
         Caption         =   "General Ledger"
      End
      Begin VB.Menu mnuTBalance 
         Caption         =   "Trial Balance"
      End
      Begin VB.Menu mnuPLAccounts 
         Caption         =   "Profit And Loss Accounts"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuRReport 
         Caption         =   "Purchase Report Information"
      End
      Begin VB.Menu mnuSPStatement 
         Caption         =   "Store wise Purchase Statement"
      End
      Begin VB.Menu mnuDReport 
         Caption         =   "Delivery Report Information"
      End
      Begin VB.Menu mnuSDStatement 
         Caption         =   "Store Wise Delivery Statement"
      End
      Begin VB.Menu mnuCustomerReport 
         Caption         =   "Department Wise Report Statement"
      End
      Begin VB.Menu mnuSposition 
         Caption         =   "Stock Status Rport"
         Begin VB.Menu mnuSSDetails 
            Caption         =   "Stock Status Detais"
         End
         Begin VB.Menu mnuSSSummary 
            Caption         =   "Stock Status Summary"
         End
      End
      Begin VB.Menu mnuMTSheet 
         Caption         =   "Monthly Top Sheet"
      End
      Begin VB.Menu MNuTopDetails 
         Caption         =   "Monthly Top Sheet Details"
      End
      Begin VB.Menu mnuSLStatement 
         Caption         =   "Supplier Ledger Statement"
      End
      Begin VB.Menu mnuSRStatement 
         Caption         =   "Supplier Report Statement"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuCommunication 
         Caption         =   "Communication"
      End
   End
   Begin VB.Menu mnuBackUp 
      Caption         =   "BackUp"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
   Begin VB.Menu mnuLogoff 
      Caption         =   "Lo&g off"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
Me.StatusBar1.Panels(2) = frmLogin.txtUID
Me.StatusBar1.Panels(3) = Date
Me.StatusBar1.Panels(4) = Time
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call mnuLogoff_Click
End Sub

Private Sub mnuBackUp_Click()
frmBackUp.Show vbModal
End Sub

Private Sub mnuCalculator_Click()
On Error Resume Next
   Shell "calc.exe"
End Sub

Private Sub mnuCBook_Click()
RptCashBook.Show vbModal
End Sub

Private Sub mnuCustomer_Click()
frmCustomer.Show vbModal
End Sub

Private Sub mnuCustomerReport_Click()
RptCustomer.Show vbModal
End Sub

Private Sub mnuDelivery_Click()
frmDelivery.Show vbModal
End Sub

Private Sub mnuDepartmentName_Click()
frmDepartment.Show vbModal
End Sub

Private Sub mnuDReport_Click()
RptDeliveryItem.Show vbModal
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuFloorsheet_Click()
RptFloorSheet.Show vbModal
End Sub

Private Sub mnuGLedger_Click()
RptGLedger.Show vbModal
End Sub

Private Sub mnuHAccounts_Click()
frmAccountsHead.Show vbModal
End Sub

Private Sub mnuHelp_Click()
FrmAbout.Show vbModal
End Sub

Private Sub mnuItemRequisation_Click()
RptReqItem.Show vbModal
End Sub

Private Sub mnuLogoff_Click()
End
End Sub

Private Sub mnuMTSheet_Click()
RptMTopSheet.Show vbModal
End Sub

Private Sub mnuProductCatagory_Click()
frmCatagory.Show vbModal
End Sub

Private Sub mnuProductSubCatagory_Click()
frmCatagorySub.Show vbModal
End Sub

Private Sub mnuReceiving_Click()
frmStock.Show vbModal
End Sub

Private Sub mnuRequisition_Click()
frmRequisition.Show vbModal
End Sub

Private Sub mnuRReport_Click()
RptReceivingItem.Show vbModal
End Sub

Private Sub mnuSDStatement_Click()
RptSDStatement.Show vbModal
End Sub

Private Sub mnuSLStatement_Click()
RptSLStatement.Show vbModal
End Sub

Private Sub mnuSname_Click()
frmSuppliername.Show vbModal
End Sub

Private Sub mnuSPInformation_Click()
frmMoneyReceipt.Show vbModal
End Sub

'Private Sub mnuSposition_Click()
'RPTStockPosition.Show vbModal
'End Sub

Private Sub mnuSPStatement_Click()
RptSRStatement.Show vbModal
End Sub

Private Sub mnuSRStatement_Click()
RptSupplier.Show vbModal
End Sub

Private Sub mnuSSDetails_Click()
RPTStockPosition.Show vbModal
End Sub

Private Sub mnuStoreName_Click()
frmStore.Show vbModal
End Sub

Private Sub MNuTopDetails_Click()
Rpt_Top_Details.Show vbModal
End Sub

Private Sub mnuUser_Click()
frmUser.Show vbModal
End Sub

Private Sub mnuVEntry_Click()
frmMoneyReceipt.Show vbModal
End Sub
