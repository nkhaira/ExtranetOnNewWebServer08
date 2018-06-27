Attribute VB_Name = "modGeneral_Routines"
'Characters
Public Const EMPTY_STRING = ""
Public Const vbSpace = " "

'-------------------------------------------------------------------------
'Purpose:   Assures that the passed path has a "\" at the end of it
'IN:
'   [sPath]
'           a valid path name
'Return:    the same path with a "\" on the end if it did not already
'           have one.
'-------------------------------------------------------------------------
Public Function FormatPath(sPath As String) As String
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    FormatPath = sPath
End Function

