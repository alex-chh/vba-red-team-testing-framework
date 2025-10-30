Attribute VB_Name = "UtilityFunctions"
' =============================================================================
' UTILITY FUNCTIONS MODULE - INTERNAL RED TEAM TESTING
' =============================================================================
' Various utility functions for system interaction, file operations, and
' environment manipulation for security testing purposes
' =============================================================================

Option Explicit

' API Declarations for advanced functionality
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetUserNameA Lib "advapi32" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare PtrSafe Function GetComputerNameA Lib "kernel32" _
    (ByVal lpBuffer As String, ByRef nSize As Long) As Long

' =============================================================================
' SYSTEM INFORMATION GATHERING
' =============================================================================

Public Function GetFullSystemInfo() As String
    ' Comprehensive system information collection
    Dim info As String
    
    info = "=== SYSTEM RECONNAISSANCE ===" & vbCrLf & _
           "Username: " & GetCurrentUserName() & vbCrLf & _
           "Computer Name: " & GetCurrentComputerName() & vbCrLf & _
           "Process ID: " & GetCurrentProcessId() & vbCrLf & _
           "OS Version: " & Environ("OS") & vbCrLf & _
           "Processor: " & Environ("PROCESSOR_IDENTIFIER") & vbCrLf & _
           "Architecture: " & Environ("PROCESSOR_ARCHITECTURE") & vbCrLf & _
           "Number of Processors: " & Environ("NUMBER_OF_PROCESSORS") & vbCrLf & _
           "User Domain: " & Environ("USERDOMAIN") & vbCrLf & _
           "User Profile: " & Environ("USERPROFILE") & vbCrLf & _
           "Temp Directory: " & Environ("TEMP") & vbCrLf & _
           "System Root: " & Environ("SystemRoot") & vbCrLf & _
           "Current Directory: " & CurDir() & vbCrLf & _
           "Office Version: " & Application.Version & vbCrLf & _
           "Word Version: " & Application.Build & vbCrLf & _
           "Execution Time: " & Now() & vbCrLf & _
           "================================"
    
    GetFullSystemInfo = info
End Function

Public Function GetCurrentUserName() As String
    ' Get current username using API call
    Dim buffer As String * 255
    Dim length As Long
    
    length = 255
    If GetUserNameA(buffer, length) Then
        GetCurrentUserName = Left$(buffer, length - 1)
    Else
        GetCurrentUserName = Environ("USERNAME")
    End If
End Function

Public Function GetCurrentComputerName() As String
    ' Get computer name using API call
    Dim buffer As String * 255
    Dim length As Long
    
    length = 255
    If GetComputerNameA(buffer, length) Then
        GetCurrentComputerName = Left$(buffer, length)
    Else
        GetCurrentComputerName = Environ("COMPUTERNAME")
    End If
End Function

' =============================================================================
' FILE SYSTEM OPERATIONS
' =============================================================================

Public Function FileExists(ByVal filePath As String) As Boolean
    ' Check if file exists
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
End Function

Public Function ReadFileContents(ByVal filePath As String) As String
    ' Read contents of a text file
    On Error GoTo ErrorHandler
    
    Dim fileNumber As Integer
    Dim content As String
    
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
    content = Input$(LOF(fileNumber), fileNumber)
    Close fileNumber
    
    ReadFileContents = content
    Exit Function
    
ErrorHandler:
    ReadFileContents = "Error reading file: " & Err.Description
End Function

Public Sub WriteToFile(ByVal filePath As String, ByVal content As String)
    ' Write content to a text file
    On Error GoTo ErrorHandler
    
    Dim fileNumber As Integer
    
    fileNumber = FreeFile
    Open filePath For Output As fileNumber
    Print #fileNumber, content
    Close fileNumber
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error writing to file: " & Err.Description
End Sub

Public Function ListFilesInDirectory(ByVal directoryPath As String, Optional ByVal pattern As String = "*.*") As Collection
    ' List files in directory matching pattern
    On Error GoTo ErrorHandler
    
    Dim fileCollection As New Collection
    Dim fileName As String
    
    If Right(directoryPath, 1) <> "\" Then directoryPath = directoryPath & "\"
    
    fileName = Dir(directoryPath & pattern)
    
    Do While fileName <> ""
        fileCollection.Add directoryPath & fileName
        fileName = Dir()
    Loop
    
    Set ListFilesInDirectory = fileCollection
    Exit Function
    
ErrorHandler:
    Set ListFilesInDirectory = New Collection
End Function

' =============================================================================
' PROCESS AND SYSTEM INTERACTION
' =============================================================================

Public Function ExecutePowerShell(ByVal command As String) As String
    ' Execute PowerShell command and return output
    On Error GoTo ErrorHandler
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    
    ' Use encoded command to avoid quoting issues
    Dim encodedCommand As String
    encodedCommand = "powershell.exe -ExecutionPolicy Bypass -NoProfile -Command " & _
                     "& { " & command & " }"
    
    ExecutePowerShell = wsh.Exec(encodedCommand).StdOut.ReadAll
    
    Exit Function
    
ErrorHandler:
    ExecutePowerShell = "PowerShell Error: " & Err.Description
End Function

Public Function GetRunningProcesses() As String
    ' Get list of running processes
    Dim result As String
    result = ExecutePowerShell("Get-Process | Select-Object Name, Id, CPU, WorkingSet | Format-Table -AutoSize")
    GetRunningProcesses = result
End Function

Public Function GetNetworkConnections() As String
    ' Get active network connections
    Dim result As String
    result = ExecutePowerShell("Get-NetTCPConnection | Where-Object {$_.State -eq 'Established'} | Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, State | Format-Table -AutoSize")
    GetNetworkConnections = result
End Function

' =============================================================================
' ENCRYPTION AND ENCODING UTILITIES
' =============================================================================

Public Function Base64Encode(ByVal text As String) As String
    ' Base64 encode string
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.nodeTypedValue = StrToBin(text)
    
    Base64Encode = xmlNode.text
End Function

Public Function Base64Decode(ByVal base64Text As String) As String
    ' Base64 decode string
    Dim xmlDoc As Object
    Dim xmlNode As Object
    
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    Set xmlNode = xmlDoc.createElement("b64")
    
    xmlNode.DataType = "bin.base64"
    xmlNode.text = base64Text
    
    Base64Decode = BinToStr(xmlNode.nodeTypedValue)
End Function

Private Function StrToBin(ByVal text As String) As Byte()
    ' Convert string to byte array
    StrToBin = StrConv(text, vbFromUnicode)
End Function

Private Function BinToStr(ByVal bytes() As Byte) As String
    ' Convert byte array to string
    BinToStr = StrConv(bytes, vbUnicode)
End Function

' =============================================================================
' REGISTRY MANIPULATION (Use with caution in testing environments)
' =============================================================================

Public Function ReadRegistryValue(ByVal keyPath As String, ByVal valueName As String) As String
    ' Read value from Windows Registry
    On Error GoTo ErrorHandler
    
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    
    ReadRegistryValue = wshShell.RegRead(keyPath & "\" & valueName)
    
    Exit Function
    
ErrorHandler:
    ReadRegistryValue = "Registry read error: " & Err.Description
End Function

Public Sub WriteRegistryValue(ByVal keyPath As String, ByVal valueName As String, ByVal valueData As String)
    ' Write value to Windows Registry
    On Error GoTo ErrorHandler
    
    Dim wshShell As Object
    Set wshShell = CreateObject("WScript.Shell")
    
    wshShell.RegWrite keyPath & "\" & valueName, valueData
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Registry write error: " & Err.Description
End Sub

' =============================================================================
' URL ENCODING UTILITIES
' =============================================================================

Public Function URLEncode(ByVal text As String) As String
    ' URL encode a string (percent encoding)
    Dim i As Integer
    Dim charCode As Integer
    Dim result As String
    
    result = ""
    
    For i = 1 To Len(text)
        charCode = Asc(Mid(text, i, 1))
        
        ' Keep alphanumeric and some safe characters as-is
        If (charCode >= 48 And charCode <= 57) Or _
           (charCode >= 65 And charCode <= 90) Or _
           (charCode >= 97 And charCode <= 122) Or _
           charCode = 45 Or charCode = 95 Or charCode = 46 Or charCode = 126 Then
            result = result & Chr(charCode)
        Else
            ' Percent encode everything else
            result = result & "%" & Right("0" & Hex(charCode), 2)
        End If
    Next i
    
    URLEncode = result
End Function
