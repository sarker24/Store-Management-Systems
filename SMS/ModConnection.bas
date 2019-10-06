Attribute VB_Name = "ModConnection"
Option Explicit
Public cn As ADODB.Connection
Public gstrConnection               As String
Public sServerName                  As String
Public SDatabaseName                As String
Public SSecurityDatabaseName        As String
Public ConStr As String
Public strCallingForm As String

'-----------------------------------
'Add for date purpose
Public rsServerDate As ADODB.Recordset
'Set rsServerDate = New ADODB.Recordset
'rsServerDate.Open "Select getDate()", cn, adOpenStatic, adLockReadOnly

Public Type POINTAPI
        x As Long
        y As Long
End Type
'----------------------------------
Public Enum USER_MODE
    VIEW_MODE = 0
    INSERT_MODE = 1
    UPDATE_MODE = 2
End Enum
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Add for date purpose
Public Declare Function ClientToScreen Lib "User32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
'--------------------------------------

Public Function Connect() As Boolean
    On Error GoTo ProcError
    
    Dim CheckConnection             As Integer
'        Dim cn                          As ADODB.Connection
        GetServer
        Set cn = New ADODB.Connection
'        gstrConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & SDatabaseName & ";Data Source=" & sServerName
'        ConStr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & SDatabaseName & ";Data Source=" & sServerName

        gstrConnection = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=" & SDatabaseName & ";Server=" & sServerName
        ConStr = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=" & SDatabaseName & ";Server=" & sServerName
        
        cn.Open gstrConnection
        CheckConnection = 1
        Connect = True
        Exit Function
ProcError:
    Select Case Err.Number
    Case 0
    Case -2147467259
        MsgBox "Your Server Is Not Available"
'        frmSeverSetup.Show vbModal
        If CheckConnection = 1 Then Connect = True
        Exit Function
    Case Else
        MsgBox Err.Description
    End Select
End Function

Private Sub GetServer()
    
    Dim fileName            As String
    Dim ApplicationName     As String
    Dim KeyName             As String
    Dim MyDatabaseName As String
'    Dim MySecurityDatabase As String
    fileName = App.Path + "\RMS.ini"
    ApplicationName = "INI"
    KeyName = "ServerName"
    MyDatabaseName = "DatabaseName"
'    MySecurityDatabase = "SecurityDatabaseName"
    Dim buf As String * 256
    
    Dim BufDatabase As String * 256
'    Dim BufSecurityDatabase As String * 256
    
    Dim length As Long
    
    Dim DBLength As Long
'    Dim DBSecurityLength As Long
    
    length = GetPrivateProfileString( _
    ApplicationName, KeyName, "<no value>", _
    buf, Len(buf), fileName)
    'Retrieve database length
    
    DBLength = GetPrivateProfileString( _
    ApplicationName, MyDatabaseName, "<no value>", _
    BufDatabase, Len(BufDatabase), fileName)
    
    'Retrieve Security Database length
'   DBSecurityLength = GetPrivateProfileString( _
'    ApplicationName, MySecurityDatabase, "<no value>", _
'    BufSecurityDatabase, Len(BufSecurityDatabase), fileName)
    
    
    
    sServerName = Strings.Left$(buf, length)
    
    'Retrieve Database Name
    
    SDatabaseName = Strings.Left$(BufDatabase, DBLength)
'    SSecurityDatabaseName = Strings.Left$(BufSecurityDatabase, DBSecurityLength)
    

End Sub

Public Function StrSecurityConnect()
    StrSecurityConnect = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & SSecurityDatabaseName & ";Data Source=" & sServerName
End Function


Public Sub MakeSound()
    
    Dim i As Integer
    For i = 1 To 10
        Beep
    Next i
    
End Sub


Sub Gitna(Frm As Form)
    Frm.Left = (frmMain.ScaleWidth - Frm.Width) / 2
    Frm.Top = (frmMain.ScaleHeight - Frm.Height) / 2
End Sub
