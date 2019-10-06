VERSION 5.00
Begin VB.Form frmSendmessage 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Message Send to Network Computer"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   Icon            =   "frmSendmessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4980
   Begin VB.CommandButton cmdSendMessage 
      Caption         =   "&Send Message"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtComputerName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtEnterMessage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label lblComputerName 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Computer Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblEnterMessage 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Enter Message"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "frmSendmessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
KeyPreview = True
Me.Left = 0
Me.Top = 700
End Sub

Private Sub cmdSendMessage_Click()
If Len(txtComputerName.text) = 0 Then
    MsgBox "Enter Computer name ...", vbInformation
    Exit Sub
End If
If Len(txtEnterMessage.text) = 0 Then
    MsgBox "Enter your message ...", vbInformation
    Exit Sub
End If
Shell "net send " & txtComputerName.text & " " & txtEnterMessage.text & vbCrLf & vbCrLf & "[ Message Sent using Browse MIS System 1.0.0 (Developed By MAS IT SOLUTIONS) ]"
MsgBox "Your message sent to " & txtComputerName.text, vbInformation
txtEnterMessage.text = Clear
End Sub

Private Sub txtComputerName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys "{TAB}"
End If
End Sub






