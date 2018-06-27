Attribute VB_Name = "modDebugUtilities"
Public Function ErrorReport(errobject As errobject, strMessage As String)
    Dim fs As Scripting.FileSystemObject
    Dim textFile As Scripting.TextStream
    Dim strDate As String
    Dim strOutputPath As String
    Dim strINIFile As String
    Dim strErrorReport As String
    
    ' Get today's date
    strDate = Month(Now) & "-" & Day(Now) & "-" & Year(Now)
    
    ' Get output path - saved in ini
    strOutputPath = App.Path
    
    Set fs = New FileSystemObject
    
    strErrorReport = strMessage & vbCrLf
    strErrorReport = strErrorReport & "Error Number: " & errobject.Number & vbCrLf
    strErrorReport = strErrorReport & "Error Description: " & errobject.Description & vbCrLf
    strErrorReport = strErrorReport & "Error Source: " & errobject.Source & vbCrLf & vbCrLf
    
    ' Save error report
    Set textFile = fs.OpenTextFile(strOutputPath & "\" & " Errors " & strDate & ".txt", ForAppending, True, TristateUseDefault)

    textFile.WriteLine strErrorReport
    textFile.Close
    
    ' Cleanup
    Set textFile = Nothing
    
    Err.Clear
End Function

Public Function RaiseError(errobject As errobject, strSource As String, strMessage As String) As Error
    Err.Raise errobject.Number, strSource, errobject.Description & vbCrLf & strMessage
End Function
