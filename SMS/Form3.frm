VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   10050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15030
   LinkTopic       =   "Form3"
   ScaleHeight     =   10050
   ScaleWidth      =   15030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delivery Master Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1575
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   14535
      Begin VB.TextBox txtRNo 
         Height          =   375
         Left            =   4560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   720
         Width           =   2535
      End
      Begin VB.ComboBox cboDName 
         Height          =   315
         Left            =   7200
         TabIndex        =   15
         Text            =   " "
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton CmdPost 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Post"
         Height          =   855
         Index           =   0
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Item Input"
         Height          =   855
         Left            =   12360
         Picture         =   "Form3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58392577
         CurrentDate     =   39646
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   58392577
         CurrentDate     =   39647
      End
      Begin VB.Label lblDDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Delivery Date"
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
         Left            =   2400
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblRDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Requisition Date"
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
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblRNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   " Requisiton No"
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
         Left            =   4560
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblDepartment 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Department Name"
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
         Left            =   7200
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delivery Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   7335
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   14535
      Begin VSFlex7DAOCtl.VSFlexGrid VSFlexGrid1 
         Height          =   6495
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   14055
         _cx             =   24791
         _cy             =   11456
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
         BackColorFixed  =   -2147483630
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Form3.frx":08CA
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.CommandButton cmdUndoPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Undo Post"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      Picture         =   "Form3.frx":09B2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9120
      Width           =   1095
   End
   Begin VB.CommandButton CmdPost 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Post"
      Height          =   855
      Index           =   1
      Left            =   6960
      Picture         =   "Form3.frx":127C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   855
      Left            =   240
      Picture         =   "Form3.frx":1B46
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   855
      Left            =   1200
      Picture         =   "Form3.frx":2410
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      Height          =   855
      Left            =   2160
      Picture         =   "Form3.frx":2CDA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   855
      Left            =   3120
      Picture         =   "Form3.frx":3264
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Q&uit"
      Height          =   855
      Left            =   4080
      Picture         =   "Form3.frx":3B2E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton chameleonButton1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      Height          =   855
      Left            =   5040
      Picture         =   "Form3.frx":43F8
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   855
      Left            =   6000
      Picture         =   "Form3.frx":4CC2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9120
      Width           =   990
   End
   Begin VB.CommandButton cmdRequisition 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Requisition"
      Height          =   855
      Left            =   9000
      Picture         =   "Form3.frx":558C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9120
      Width           =   990
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
