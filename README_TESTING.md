# VBA Red Team Testing Framework

## Overview
This framework provides VBA modules for internal red team testing and security assessment purposes. The modules demonstrate various techniques for command and control (C2) communication, system reconnaissance, and evasion mechanisms.

## ⚠️ IMPORTANT LEGAL DISCLAIMER
**FOR INTERNAL TESTING ONLY**
- Use only in authorized testing environments
- Ensure proper written authorization before deployment
- Do not use against systems without explicit permission
- Comply with all applicable laws and regulations

## Modules Created

### 1. C2MainModule.bas
Primary command and control functionality:
- System information gathering (GatherSystemInformation)
- Process execution via cmd.exe (ExecuteCommand)
- Edge browser-based C2 channel establishment (EstablishEdgeC2)
- HTTP beaconing to C2 server (SendBeacon)
- C2 connection initialization (InitializeC2Connection)

### 2. UtilityFunctions.bas
Comprehensive utility functions:
- System reconnaissance (GetFullSystemInfo, GetCurrentUserName, GetCurrentComputerName)
- File system operations (FileExists, ReadFileContents, ListFilesInDirectory)
- PowerShell command execution (ExecutePowerShell)
- Process and network enumeration (GetRunningProcesses, GetNetworkConnections)
- Encoding/decoding (Base64Encode, Base64Decode, URLEncode)
- Registry operations (ReadRegistryValue)

### 3. ObfuscationTechniques.bas
Anti-analysis and evasion techniques:
- String obfuscation (ObfuscateString, DeobfuscateString)
- Environment detection (IsDebuggerPresent, IsSandbox)
- Polymorphic code execution (ExecutePolymorphicCode, GenerateDynamicCode)
- Artifact cleaning (CleanArtifacts)
- Dynamic function calls (DynamicFunctionCall)
- Anti-analysis checks (AntiAnalysisChecks)

### 4. AutoExecModule.bas
Automatic execution and persistence:
- Auto-execution macros (AutoExec, AutoOpen, Document_Open)
- Delayed execution (DelayedAutoExecution, DelayedDocumentOpen)
- Stealth C2 initialization (InitializeSilentC2, ExecuteStealthC2)
- Environment safety checks (IsSafeEnvironment, IsAnalysisToolRunning)
- Persistence mechanisms (EstablishPersistence)
- Process existence checking (ProcessExists)

### 5. StartupModule.bas
Document startup and stealth execution:
- Document event handlers (AutoOpen, Document_Open)
- Delayed startup execution (DelayedStartup)
- Environment detection (IsEnvironmentSafe, IsDebuggerPresent, IsSandboxEnvironment)
- Stealth C2 execution (ExecuteStealthC2)
- System information beaconing (SendSystemInfoBeacon)
- Low-level API declarations for advanced functionality

## Installation and Usage

### Importing into Word
1. Open Microsoft Word
2. Press `ALT + F11` to open VBA Editor
3. Right-click on "Normal" project → Import File
4. Select all five `.bas` files from this directory:
   - C2MainModule.bas
   - UtilityFunctions.bas
   - ObfuscationTechniques.bas
   - AutoExecModule.bas
   - StartupModule.bas

### Basic Testing Commands
```vba
' === C2MainModule ===
InitializeC2Connection           ' Initialize C2 connection
EstablishEdgeC2                 ' Establish Edge browser C2
Dim result As String
result = ExecuteCommand("whoami") ' Execute system command
Debug.Print result

' === UtilityFunctions ===
Dim sysInfo As String
sysInfo = GetFullSystemInfo()     ' Get comprehensive system info
Debug.Print sysInfo

Dim processes As String
processes = GetRunningProcesses() ' Get running processes
Debug.Print processes

' === ObfuscationTechniques ===
MainObfuscated                  ' Test obfuscation techniques

Dim obfuscated As String
obfuscated = ObfuscateString("secret data") ' String obfuscation
Debug.Print "Obfuscated: " & obfuscated
Debug.Print "Deobfuscated: " & DeobfuscateString(obfuscated)

' === AutoExecModule === (Auto-executes when document opens)
IsSafeEnvironment()             ' Check environment safety

' === StartupModule === (Auto-executes when document opens)
IsEnvironmentSafe()             ' Environment safety check
```

### Advanced Testing Commands
```vba
' Test environment detection
If ObfuscationTechniques.IsDebuggerPresent() Then
    Debug.Print "Debugger detected"
Else
    Debug.Print "No debugger present"
End If

If ObfuscationTechniques.IsSandbox() Then
    Debug.Print "Sandbox environment detected"
Else
    Debug.Print "Normal environment"
End If

' Test polymorphic behavior
ObfuscationTechniques.ExecutePolymorphicCode

' Test file operations
If FileExists("C:\\Windows\\System32\\cmd.exe") Then
    Debug.Print "CMD.exe exists"
End If

' Test network connections
Dim connections As String
connections = GetNetworkConnections()
Debug.Print connections
```

## Testing Scenarios

### 1. Detection Evasion Testing
- Execute obfuscated code to test AV/EDR detection
- Test sandbox and debugger detection capabilities
- Evaluate polymorphic behavior detection

### 2. C2 Communication Testing
- Test beaconing behavior detection
- Evaluate network traffic analysis
- Test command execution monitoring

### 3. Persistence Testing
- Test registry modification detection
- Evaluate file system artifact detection
- Test process creation monitoring

## Security Considerations

### Defensive Measures to Test
1. **Macro Security**
   - Test against various macro security levels
   - Evaluate digitally signed macro detection
   - Test Office security policy enforcement

2. **Process Monitoring**
   - Test detection of child process creation
   - Evaluate command line argument monitoring
   - Test network connection detection

3. **Memory Analysis**
   - Test in-memory execution detection
   - Evaluate API call monitoring
   - Test code injection detection

### Detection Bypass Techniques
- String obfuscation for signature evasion
- Environmental awareness for sandbox evasion
- Timing-based evasion techniques
- Legitimate process abuse (living off the land)

## Monitoring and Logging

### Recommended Log Sources
1. **Windows Event Logs**
   - Security event ID 4688 (process creation)
   - PowerShell operational logs
   - Office macro execution events

2. **Network Monitoring**
   - DNS queries to C2 domains
   - HTTP/HTTPS traffic patterns
   - Certificate validation anomalies

3. **Endpoint Detection**
   - Process tree anomalies
   - File creation in temp directories
   - Registry modifications

## Cleanup Procedures

### Post-Testing Cleanup
```vba
' Clean artifacts
CleanArtifacts

' Remove registry modifications (if any)
' Note: Manual review recommended for registry changes
```

### System Restoration
1. Close any Edge instances opened during testing
2. Clear temporary files
3. Restore original macro security settings
4. Remove any test registry entries

## Compliance and Documentation

### Required Documentation
- Written authorization for testing
- Test scope and boundaries
- Incident response plan
- Data handling procedures

### Reporting
- Document all test activities
- Record detection/prevention events
- Note any false positives/negatives
- Provide recommendations for improvement

## Troubleshooting

### Common Issues
1. **Macro Security Warnings**
   - Ensure testing environment has appropriate security settings
   - Use digitally signed macros for testing

2. **Permission Errors**
   - Run Word as administrator if registry access needed
   - Ensure proper user permissions for testing

3. **Detection Triggers**
   - Adjust obfuscation techniques if detected
   - Modify timing and patterns to avoid detection

## Support
For issues with this testing framework, contact your red team lead or security operations team.

---

**Remember**: Always test responsibly and within authorized boundaries. Unauthorized testing may violate laws and organizational policies.
