Attribute VB_Name = "modSupport_ServiceSubscription"
Option Explicit

' User-defined values
Private INI_FILE
Private INI_APP

Public Function RemoveExtraInfo(strOutput As String) As String
    Dim iPosStart As Integer
    Dim iPosEnd As Integer
    Dim iStart As Integer
    
    iStart = 1
    Do While InStr(iStart, strOutput, "[") > 0
        iPosStart = InStr(1, strOutput, "[")
        iPosEnd = InStr(iPosStart, strOutput, "]")
        strOutput = Mid(strOutput, 1, iPosStart - 1) & Mid(strOutput, iPosEnd + 1)
        iStart = iPosStart + 1
    Loop
    RemoveExtraInfo = strOutput
End Function

Public Function SQLCleanup(strFieldValue As String) As String
    Dim strOutput As String
    
    strOutput = strFieldValue
    strOutput = Replace(strOutput, "[", vbNullString)
    strOutput = Replace(strOutput, "]", vbNullString)
    strOutput = Replace(strOutput, ",", vbNullString)
    ' Have to do this hack because replace function has trouble with vbcrlf when there is
    ' a chr(10), chr(13), and then a chr(10) again. It only replaces instances where both appear
    strOutput = Replace(strOutput, vbCrLf, vbNullString)
    strOutput = Replace(strOutput, Chr(10), vbNullString)
    strOutput = Replace(strOutput, Chr(13), vbNullString)
    SQLCleanup = strOutput
End Function
