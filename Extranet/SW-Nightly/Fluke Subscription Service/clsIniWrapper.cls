VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API declarations
Private Declare Function GetPrivateProfileString& Lib "Kernel32" Alias "GetPrivateProfileStringA" _
                    ( _
                    ByVal AppName As String, _
                    ByVal KeyName As String, _
                    ByVal keydefault As String, _
                    ByVal ReturnString As String, _
                    ByVal NumBytes As Integer, _
                    ByVal FileName As String _
                    )

' User defined values
Private g_strINI_FILE As String
Private g_strINI_APP As String

Private Const EMPTY_STRING = ""

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

Public Function GetIniString(strKey As String) As String
    Dim strOutput As String
    
    strOutput = VBGetPrivateProfileString(g_strINI_APP, strKey, EMPTY_STRING, g_strINI_FILE)
    
    GetIniString = strOutput
End Function

'-------------------------------------------------------------------------
'Purpose:   Get a string from a private INI file
'In:
' [sAppName] Section name in INI file
' [sKeyName] Key name in Section
' [sDefault] Default return value if key not found
' [sIniFile] Name of INI file to access
'Return:
'   Value of Key is returned, or "" if not found.
'   NOTE: Maximum length of returned string is 100.
'-------------------------------------------------------------------------
Private Function VBGetPrivateProfileString(sAppName As String, _
                                          sKeyName As String, _
                                          sDefault As String, _
                                          sINIFile As String) As String
    Const MAX_STR_LEN = 100
    Dim sProfileBuf As String
    Dim lReturn        As Long

    ' fill buffer with nulls
    sProfileBuf = String$(MAX_STR_LEN, 0)
    lReturn = GetPrivateProfileString(sAppName, _
                                        ByVal sKeyName, _
                                        sDefault, _
                                        sProfileBuf, _
                                        MAX_STR_LEN, _
                                        sINIFile)
    If lReturn Then
        lReturn = InStr(1, sProfileBuf, vbNullChar) - 1        ' look for first null
        VBGetPrivateProfileString = Left$(sProfileBuf, lReturn)
    Else
        VBGetPrivateProfileString = EMPTY_STRING
    End If
End Function
