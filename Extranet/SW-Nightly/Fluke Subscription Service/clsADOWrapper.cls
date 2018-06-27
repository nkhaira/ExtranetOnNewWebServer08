VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsADOWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Constants
Private Const CONNECTION_TIMEOUT = 4 'seconds
Private Const COMMAND_TIMEOUT = 9 'seconds
Private Const DB_SQLSERVER_CONNECT_STRING = "DRIVER=SQL Server;SERVER=[SERVER];UID=[UID];PWD=[PWD];APP=Microsoft Development Environment;WSID=;DATABASE=[DATABASE];"
Private Const DB_MSACCESS_CONNECT_STRING = "DSN=[DSN];UID=[UID];PWD=[PWD];FIL=MS Access;MaxBuefferSize=2048;PageTimeout=5;"

'INI file constants
Private Const INI_KEY_UID = "UID"
Private Const INI_KEY_PWD = "PWD"
Private Const INI_KEY_SERVER = "SERVER"
Private Const INI_KEY_DATABASE = "DATABASE"
Private Const INI_KEY_SENDEREMAIL = "SENDEREMAIL"
Private Const INI_KEY_RECEIVEREMAIL = "RECEIVEREMAIL"
Private Const INI_KEY_MAILSERVER = "MAILSERVER"
Private Const INI_OUTPUT_PATH = "OutputPath"
Private Const INI_KEY_RECORDSEED = "RECORDSEED"
Private Const INI_KEY_STARTDATE = "STARTDATE"
Private Const INI_KEY_ENDDATE = "ENDDATE"
Private Const INI_KEY_DSN = "DSN"

Private Const TOKEN_UID = "[UID]"
Private Const TOKEN_PWD = "[PWD]"
Private Const TOKEN_SERVER = "[SERVER]"
Private Const TOKEN_DATABASE = "[DATABASE]"
Private Const TOKEN_DSN = "[DSN]"

' User defined values
Private g_strINI_FILE
Private g_strINI_APP

Private Const EMPTRY_STRING = ""

Public Enum eDatabaseType
    SQLServer = 0
    MSAccess = 1
End Enum

Public Enum eCommandType
    StoredProc = 0
    SQLString = 1
End Enum

Public Property Get INI_FILE() As String
    INI_FILE = g_strINI_FILE
End Property

Public Property Let INI_FILE(ByVal strINI_FILE As String)
    g_strINI_FILE = strINI_FILE
End Property

Public Property Get INI_APP() As String
    INI_APP = g_strINI_APP
End Property

Public Property Let INI_APP(ByVal strINI_APP As String)
    g_strINI_APP = strINI_APP
End Property

'---------------------------------------------------------------
'Purpose:   Gets an ADO Connection to the database
'---------------------------------------------------------------
Public Function GetConnection(ByVal eDatabaseType As eDatabaseType) As ADODB.Connection
    On Error GoTo EXIT_GetConnection
    
    Dim oConnection As ADODB.Connection
    Dim sConnectString As String

    Set oConnection = New ADODB.Connection
    With oConnection
        .ConnectionTimeout = CONNECTION_TIMEOUT
        .CommandTimeout = COMMAND_TIMEOUT
        .CursorLocation = adUseClient
        .Mode = adModeReadWrite
    End With
    
    oConnection.Open GetConnectionString(eDatabaseType)

    Set GetConnection = oConnection
Exit Function
    
EXIT_GetConnection:
    RaiseError err, "clsADOWrapper.GetConnection()", "Could not establish a connection to the database."
End Function

Public Function CloseConnection(adoConnection As ADODB.Connection)
    On Error GoTo EXIT_CloseConnection
    
    adoConnection.Close
    Set adconnection = Nothing
    
EXIT_CloseConnection:
    RaiseError err, "clsADOWrapper.CloseConnection()", "Could not close database connection."
    'ErrorReport Err, "Error establishing connection to database", strOutputType
End Function

Private Function GetConnectionString(ByVal eDatabaseType As eDatabaseType) As String
    On Error GoTo EXIT_GetConnectionString
    
    Dim oIni As New clsIniWrapper
    Dim strConnect As String
    Dim str_INIFile As String

    If eDatabaseType = SQLServer Then
        strConnect = DB_SQLSERVER_CONNECT_STRING
    Else
        strConnect = DB_MSACCESS_CONNECT_STRING
    End If
    
    Set oIni = New clsIniWrapper
    With oIni
        .INI_FILE = g_strINI_FILE
        .INI_APP = g_strINI_APP
        strConnect = Replace(strConnect, TOKEN_SERVER, .GetIniString(INI_KEY_SERVER))
        strConnect = Replace(strConnect, TOKEN_DSN, .GetIniString(INI_KEY_DSN))
        strConnect = Replace(strConnect, TOKEN_DATABASE, .GetIniString(INI_KEY_DATABASE))
        strConnect = Replace(strConnect, TOKEN_UID, .GetIniString(INI_KEY_UID))
        strConnect = Replace(strConnect, TOKEN_PWD, .GetIniString(INI_KEY_PWD))
    End With
    
    Set oIni = Nothing
    
    GetConnectionString = strConnect
Exit Function

EXIT_GetConnectionString:
    RaiseError err, "clsADOWrapper.GetConnectionString()", "Could not get connection string."
End Function

Public Function GetRecordset(ByVal strCommand, ByRef oConnection As ADODB.Connection, ByVal eCommandType As eCommandType) As ADODB.Recordset
    On Error GoTo EXIT_GetRecordset

    Dim oCommand As ADODB.Command
    Dim rsRecordset As ADODB.Recordset
    
    Select Case eCommandType
        Case eCommandType = SQLString
            Set rsRecordset = oConnection.Execute(strCommand)
        Case eCommandType = StoredProc
            Set oCommand = New ADODB.Command
            With oCommand
                Set .ActiveConnection = oConnection
                .CommandType = adCmdStoredProc
                .CommandText = strCommand
                Set rsRecordset = .Execute
            End With
    End Select

    Set GetRecordset = rsRecordset

    ' Cleanup
    Set rsRecordset = Nothing
    Set oCommand = Nothing
Exit Function

EXIT_GetRecordset:
    RaiseError err, "clsADOWrapper.GetRecordset()", "Could not get recordset."
End Function
