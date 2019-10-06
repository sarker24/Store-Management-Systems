VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date Picker"
   ClientHeight    =   2310
   ClientLeft      =   4545
   ClientTop       =   2460
   ClientWidth     =   2850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   16711680
      BackColor       =   -2147483639
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   62062593
      CurrentDate     =   38015
   End
   Begin VB.Label lblRow 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label lblCol 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label4"
      Height          =   300
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim dt As Variant

Sub UpdateDate()
    If Not Enabled Then Exit Sub
    'If Not IsDate(txtMonth & "/" & txtDay & "/" & txtYear) Then Beep: Exit Sub
    dt = Format(MonthView1.Value, "dd-mmm-yyyy")
    'lblDate = Format(dt, "Long Date")
    Tag = dt
End Sub

Private Sub Form_Activate()
    dt = Tag
    Enabled = False
    MonthView1.Value = CDate(dt)
    Enabled = True
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    
    ' update grid value
    If IsDate(Tag) Then
        Select Case LCase(strCallingForm)
            Case LCase("frmStock")
                 frmStock.fgStock.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
                 frmStock.fgStock.SetFocus
'            Case LCase("frmLabdip")
'                 frmLabdip.vfgLabdipDet.Cell(flexcpText, lblRow, lblCol) = Format(Tag, "dd-mmm-yyyy")
'                 frmLabdip.vfgLabdipDet.SetFocus
      End Select
    End If
    
    ' go away
    Hide
    
    strCallingForm = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ' esc clears, then quits
    If KeyAscii = 27 Then Tag = "": Hide
    
    ' enter quits
    If KeyAscii = 13 Then
        UpdateDate
        Hide
    End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    UpdateDate
End Sub

Private Sub MonthView1_DateDblClick(ByVal DateDblClicked As Date)
    UpdateDate
    Hide
End Sub
