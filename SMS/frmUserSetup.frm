VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "User Set-Up"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "User Information Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   2400
         Picture         =   "frmUserSetup.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdOpen 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Open"
         Height          =   735
         Left            =   4320
         Picture         =   "frmUserSetup.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Edit"
         Height          =   735
         Left            =   1440
         Picture         =   "frmUserSetup.frx":0FB4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdNew 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&New"
         Height          =   735
         Left            =   480
         Picture         =   "frmUserSetup.frx":187E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Q&uit"
         Height          =   735
         Left            =   3360
         Picture         =   "frmUserSetup.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3000
         Width           =   990
      End
      Begin VB.TextBox TxtUName 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox txtPassward 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1560
         Width           =   3135
      End
      Begin VB.TextBox txtConPas 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ComboBox cboEn 
         BackColor       =   &H00D0B5A8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   2
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox txtSNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label MSG 
         Alignment       =   2  'Center
         BackColor       =   &H00D0B5A8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   600
         TabIndex        =   16
         Top             =   3960
         Width           =   4575
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblPgroup 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Privilege Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Passward "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Confirm Passward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblSerial 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'      Option Explicit
'
'      'Add reference to Crystal Reports x.x ActiveX Designer RunTime Library
'  'Add reference to Crystal Reports Viewer Control
'
'      'Add reference to Microsoft ActiveX Data Objects 2.x Library
'
'      'oCnn = current open ADO connection object
'
'      Private Sub Command1_Click()
'
'
'
'          Dim oApp As CRAXDRT.Application
'
'          Dim oReport As CRAXDRT.Report
'
'          Dim oRs As ADODB.Recordset
'
'          Dim sSQL As String
'
'
'
'          sSQL = "SELECT * FROM Table1"
'
'          Set oRs = New ADODB.Recordset
'       Set oRs = oCnn.Execute(sSQL)
'
'          Set oApp = New CRAXDRT.Application
'
'          Set oReport = oApp.OpenReport(App.Path & "\MyReport.rpt", 1)
'
'          oReport.Database.SetDataSource oRs, 3, 1
'
'          crvMyCRViewer.ReportSource = oReport
'
'          crvMyCRViewer.ViewReport
'
'
'
'      End Sub
