' ============================================================================
' MODULE: modLocalTableManager (Enhanced v2.0)
' PURPOSE: High-level API for local table cache system
' STRATEGY: Lazy loading, respect configuration, minimal startup impact
' ============================================================================
Option Compare Database
Option Explicit

' === Global Cache Instance ===
Public gLocalTableCache As clsLocalTableCache

' === Constants ===
Private Const LOG_PREFIX As String = "Manager> "

' ============================================================================
' APPLICATION LIFECYCLE
' ============================================================================

' ---------------------------------------------------------------------------
' ApplicationStartup - Minimal initialization with lazy loading
' ---------------------------------------------------------------------------
Public Sub ApplicationStartup()
    On Error GoTo ErrHandler
    
    Debug.Print String(70, "=")
    Debug.Print LOG_PREFIX & "APPLICATION STARTUP"
    Debug.Print LOG_PREFIX & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
                Environ("USERNAME") & "@" & Environ("COMPUTERNAME")
    Debug.Print String(70, "=")
    
    ' 1. Set backend type (configure as needed)
    gBackendType = Backend_Access
    Debug.Print LOG_PREFIX & "Backend: " & IIf(gBackendType = Backend_SQLServer, "SQL Server", "Access")
    
    ' 2. Initialize connection
    InitializeConnection
    Debug.Print LOG_PREFIX & "Connection initialized"
    
    ' 3. Validate system tables
    If Not ValidateSystemTables() Then
        MsgBox "Cache system configuration tables are missing or invalid." & vbCrLf & _
               "The application may not function correctly.", _
               vbExclamation, "Configuration Error"
        Exit Sub
    End If
    
    ' 4. Initialize cache object
    Set gLocalTableCache = New clsLocalTableCache
    Debug.Print LOG_PREFIX & "Cache engine initialized"
    
    ' 5. Load ONLY tables marked for startup (respects RefreshOnStartup flag)
    gLocalTableCache.LoadStartupTables
    
    ' 6. Display summary
    Debug.Print String(70, "-")
    Debug.Print LOG_PREFIX & GetCacheStatistics()
    Debug.Print String(70, "=")
    Debug.Print LOG_PREFIX & "Startup complete (lazy loading enabled)"
    Debug.Print String(70, "=")
    
    Exit Sub
    
ErrHandler:
    MsgBox "Critical startup error:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "The application may not function correctly.", _
           vbCritical, "Startup Error"
    Debug.Print LOG_PREFIX & "CRITICAL ERROR: " & Err.Description
End Sub

' ---------------------------------------------------------------------------
' ApplicationShutdown - Cleanup and optional table deletion
' ---------------------------------------------------------------------------
Public Sub ApplicationShutdown()
    On Error Resume Next
    
    Debug.Print String(70, "=")
    Debug.Print LOG_PREFIX & "APPLICATION SHUTDOWN"
    Debug.Print LOG_PREFIX & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print String(70, "=")
    
    ' Display final statistics
    Debug.Print LOG_PREFIX & GetCacheStatistics()
    
    ' Cleanup tables marked for deletion (respects DeleteOnShutdown flag)
    If Not gLocalTableCache Is Nothing Then
        Dim deletedCount As Long
        deletedCount = gLocalTableCache.CleanupOnShutdown()
        
        If deletedCount > 0 Then
            Debug.Print LOG_PREFIX & "Deleted " & deletedCount & " temporary table(s)"
        Else
            Debug.Print LOG_PREFIX & "No tables marked for deletion"
        End If
    End If
    
    ' Release objects
    Set gLocalTableCache = Nothing
    Debug.Print LOG_PREFIX & "Cache engine released"
    
    ' Close connection
    CloseGlobalConnection
    Debug.Print LOG_PREFIX & "Connection closed"
    
    ' Clear status bar
    SysCmd acSysCmdClearStatus
    
    Debug.Print String(70, "=")
    Debug.Print LOG_PREFIX & "Shutdown complete"
    Debug.Print String(70, "=")
End Sub

' ============================================================================
' CACHE ACCESS API
' ============================================================================

' ---------------------------------------------------------------------------
' EnsureTableReady - Main API: Call before using any cached table
' This triggers lazy loading if table hasn't been loaded yet
' ---------------------------------------------------------------------------
Public Function EnsureTableReady(LookupName As String) As Boolean
    On Error GoTo ErrHandler
    
    ' Validate input
    If Trim(LookupName) = "" Then
        Debug.Print LOG_PREFIX & "ERROR: Empty lookup name"
        Exit Function
    End If
    
    ' Initialize cache if needed
    If gLocalTableCache Is Nothing Then
        Debug.Print LOG_PREFIX & "Cache not initialized, initializing now..."
        Set gLocalTableCache = New clsLocalTableCache
    End If
    
    ' Ensure table is ready (loads if needed)
    EnsureTableReady = gLocalTableCache.EnsureTableReady(LookupName)
    
    Exit Function
    
ErrHandler:
    Debug.Print LOG_PREFIX & "ERROR in EnsureTableReady: " & Err.Description
    EnsureTableReady = False
End Function

' ---------------------------------------------------------------------------
' ForceRefresh - Manually refresh a specific table
' ---------------------------------------------------------------------------
Public Function ForceRefresh(LookupName As String) As Boolean
    On Error GoTo ErrHandler
    
    If Trim(LookupName) = "" Then Exit Function
    
    If gLocalTableCache Is Nothing Then
        Set gLocalTableCache = New clsLocalTableCache
    End If
    
    Debug.Print LOG_PREFIX & "Manual refresh requested: " & LookupName
    
    Dim startTime As Double
    startTime = Timer
    
    ForceRefresh = gLocalTableCache.ForceRefresh(LookupName)
    
    If ForceRefresh Then
        Debug.Print LOG_PREFIX & "Refresh completed in " & _
                    Format((Timer - startTime), "0.00") & " seconds"
        MsgBox "Table '" & LookupName & "' refreshed successfully.", vbInformation
    Else
        MsgBox "Failed to refresh table '" & LookupName & "'." & vbCrLf & _
               "Check debug output for details.", vbExclamation
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print LOG_PREFIX & "ERROR in ForceRefresh: " & Err.Description
    ForceRefresh = False
End Function

' ---------------------------------------------------------------------------
' RefreshAll - Refresh all active tables
' ---------------------------------------------------------------------------
Public Sub RefreshAll()
    On Error GoTo ErrHandler
    
    Dim response As VbMsgBoxResult
    response = MsgBox("This will refresh ALL active cache tables." & vbCrLf & vbCrLf & _
                     "This may take several minutes. Continue?", _
                     vbQuestion + vbYesNo, "Refresh All Tables")
    
    If response <> vbYes Then Exit Sub
    
    DoCmd.Hourglass True
    
    Dim rs As DAO.Recordset
    Dim successCount As Long
    Dim failCount As Long
    
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT LookupName FROM APP_LookupConfig " & _
        "WHERE IsActive=True AND IsStatic=True ORDER BY LoadOrder")
    
    Debug.Print LOG_PREFIX & "Refreshing all tables..."
    
    Do Until rs.EOF
        If ForceRefresh(rs!LookupName) Then
            successCount = successCount + 1
        Else
            failCount = failCount + 1
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    DoCmd.Hourglass False
    
    MsgBox "Refresh complete!" & vbCrLf & vbCrLf & _
           "Success: " & successCount & vbCrLf & _
           "Failed: " & failCount, vbInformation
    
    Exit Sub
    
ErrHandler:
    DoCmd.Hourglass False
    MsgBox "Error during refresh: " & Err.Description, vbCritical
End Sub

' ============================================================================
' COMBO BOX BINDING
' ============================================================================

' ---------------------------------------------------------------------------
' BindComboFromCache - Binds combo to cached table with auto-refresh
' ---------------------------------------------------------------------------
Public Function BindComboFromCache( _
    theCombo As Access.ComboBox, _
    LookupName As String, _
    Optional searchField As String = "", _
    Optional searchType As searchType = searchType.anywhereinstring, _
    Optional handleArrows As Boolean = True) As FindAsYouTypeCombo
    
    On Error GoTo ErrHandler
    
    ' Validate inputs
    If theCombo Is Nothing Then
        MsgBox "Invalid combo box reference.", vbCritical
        Exit Function
    End If
    
    If Trim(LookupName) = "" Then
        MsgBox "Lookup name is required.", vbCritical
        Exit Function
    End If
    
    ' Ensure table is ready (triggers lazy load if needed)
    If Not EnsureTableReady(LookupName) Then
        MsgBox "Failed to prepare cache table: " & LookupName & vbCrLf & vbCrLf & _
               "The combo box may not function correctly.", _
               vbExclamation, "Cache Error"
        Exit Function
    End If
    
    ' Get local table name
    Dim localTableName As String
    localTableName = GetLocalTableName(LookupName)
    
    If localTableName = "" Then
        MsgBox "Configuration error: No local table defined for " & LookupName, vbExclamation
        Exit Function
    End If
    
    ' Set combo row source
    theCombo.RowSourceType = "Table/Query"
    theCombo.RowSource = "SELECT * FROM [" & localTableName & "] ORDER BY 2"
    
    ' Initialize Find-As-You-Type
    Dim combo As New FindAsYouTypeCombo
    combo.InitalizeFilterCombo theCombo, searchField, searchType, handleArrows
    
    Set BindComboFromCache = combo
    
    Debug.Print LOG_PREFIX & "Combo bound: " & theCombo.name & " -> " & LookupName
    Exit Function
    
ErrHandler:
    MsgBox "Error binding combo box:" & vbCrLf & vbCrLf & _
           "Control: " & theCombo.name & vbCrLf & _
           "Lookup: " & LookupName & vbCrLf & vbCrLf & _
           Err.Description, vbCritical
    Debug.Print LOG_PREFIX & "ERROR in BindComboFromCache: " & Err.Description
End Function

' ============================================================================
' INFORMATION & DIAGNOSTICS
' ============================================================================

' ---------------------------------------------------------------------------
' GetCacheStatistics - Returns summary of cache state
' ---------------------------------------------------------------------------
Public Function GetCacheStatistics() As String
    On Error Resume Next
    
    Dim stats As String
    Dim rs As DAO.Recordset
    
    ' Count configurations
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT COUNT(*) AS Total FROM APP_LookupConfig WHERE IsActive=True AND IsStatic=True")
    Dim configCount As Long: configCount = Nz(rs!total, 0)
    rs.Close
    
    ' Count startup tables
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT COUNT(*) AS Total FROM APP_LookupConfig " & _
        "WHERE IsActive=True AND IsStatic=True AND RefreshOnStartup=True")
    Dim startupCount As Long: startupCount = Nz(rs!total, 0)
    rs.Close
    
    ' Count existing tables
    Dim existingCount As Long
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If Left(tdf.name, 6) = "Local_" And Not Left(tdf.name, 4) = "MSys" Then
            existingCount = existingCount + 1
        End If
    Next
    
    stats = "Configured: " & configCount & " tables" & vbCrLf & _
            "Startup: " & startupCount & " tables" & vbCrLf & _
            "Cached: " & existingCount & " tables" & vbCrLf & _
            "Strategy: Lazy loading"
    
    GetCacheStatistics = stats
End Function

' ---------------------------------------------------------------------------
' GetTableInfo - Get detailed info about specific table
' ---------------------------------------------------------------------------
Public Function GetTableInfo(LookupName As String) As String
    On Error Resume Next
    
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim info As String
    
    sql = "SELECT c.*, s.LastRefreshDate, s.LastRecordCount, s.LastRefreshStatus, " & _
          "s.LastRefreshDuration, s.LastErrorMessage " & _
          "FROM APP_LookupConfig c " & _
          "LEFT JOIN APP_LookupClientState s ON c.LookupName = s.LookupName " & _
          "AND s.ClientMachineName='" & Environ("COMPUTERNAME") & "' " & _
          "WHERE c.LookupName='" & SQLSafe(LookupName) & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.EOF Then
        GetTableInfo = "Configuration not found"
        rs.Close
        Exit Function
    End If
    
    info = "TABLE: " & LookupName & vbCrLf & _
           String(50, "-") & vbCrLf & _
           "Local Table: " & Nz(rs!LocalTableName, "N/A") & vbCrLf & _
           "Startup Load: " & IIf(Nz(rs!RefreshOnStartup, False), "Yes", "No") & vbCrLf & _
           "Lazy Load: " & IIf(Nz(rs!LazyLoad, False), "Yes", "No") & vbCrLf & _
           "Refresh Interval: " & Nz(rs!RefreshIntervalHours, 0) & " hours" & vbCrLf & _
           "Version: " & Nz(rs!RefreshVersion, 1) & vbCrLf & _
           "Delete on Shutdown: " & IIf(Nz(rs!DeleteOnShutdown, False), "Yes", "No") & vbCrLf & _
           vbCrLf & _
           "CURRENT STATE:" & vbCrLf & _
           "Last Refresh: " & Format(Nz(rs!LastRefreshDate, "Never"), "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
           "Records: " & Nz(rs!LastRecordCount, 0) & vbCrLf & _
           "Status: " & Nz(rs!LastRefreshStatus, "Not loaded") & vbCrLf & _
           "Duration: " & Nz(rs!LastRefreshDuration, 0) & " ms"
    
    If Not IsNull(rs!LastErrorMessage) And rs!LastErrorMessage <> "" Then
        info = info & vbCrLf & "Last Error: " & rs!LastErrorMessage
    End If
    
    rs.Close
    GetTableInfo = info
End Function

' ---------------------------------------------------------------------------
' ShowCacheReport - Display full cache report
' ---------------------------------------------------------------------------
Public Sub ShowCacheReport()
    On Error Resume Next
    
    Dim Report As String
    Dim rs As DAO.Recordset
    
    Report = "LOCAL TABLE CACHE REPORT" & vbCrLf & _
            String(70, "=") & vbCrLf & vbCrLf & _
            GetCacheStatistics() & vbCrLf & vbCrLf & _
            "TABLE DETAILS:" & vbCrLf & _
            String(70, "-") & vbCrLf
    
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT c.LookupName, c.LocalTableName, c.RefreshOnStartup, c.LazyLoad, " & _
        "c.DeleteOnShutdown, s.LastRefreshDate, s.LastRecordCount, s.LastRefreshStatus " & _
        "FROM APP_LookupConfig c " & _
        "LEFT JOIN APP_LookupClientState s ON c.LookupName = s.LookupName " & _
        "AND s.ClientMachineName='" & Environ("COMPUTERNAME") & "' " & _
        "WHERE c.IsActive=True AND c.IsStatic=True ORDER BY c.LookupName")
    
    Do Until rs.EOF
        Report = Report & vbCrLf & rs!LookupName & vbCrLf & _
                "  Table: " & rs!LocalTableName & vbCrLf & _
                "  Startup: " & IIf(Nz(rs!RefreshOnStartup, False), "Yes", "No") & _
                " | Lazy: " & IIf(Nz(rs!LazyLoad, False), "Yes", "No") & _
                " | Delete: " & IIf(Nz(rs!DeleteOnShutdown, False), "Yes", "No") & vbCrLf & _
                "  Records: " & Nz(rs!LastRecordCount, "N/A") & vbCrLf & _
                "  Last Refresh: " & Format(Nz(rs!LastRefreshDate, "Never"), "mm/dd/yy hh:nn") & vbCrLf & _
                "  Status: " & Nz(rs!LastRefreshStatus, "Not loaded") & vbCrLf
        rs.MoveNext
    Loop
    
    rs.Close
    
    Debug.Print Report
    MsgBox Report, vbInformation, "Cache Report"
End Sub

' ============================================================================
' VALIDATION & TESTING
' ============================================================================

' ---------------------------------------------------------------------------
' ValidateSystemTables - Checks if required tables exist
' ---------------------------------------------------------------------------
Private Function ValidateSystemTables() As Boolean
    On Error Resume Next
    
    ValidateSystemTables = True
    
    ' Check config table
    If Not TableExists("APP_LookupConfig") Then
        Debug.Print LOG_PREFIX & "ERROR: APP_LookupConfig table not found"
        ValidateSystemTables = False
    End If
    
    ' Check state table
    If Not TableExists("APP_LookupClientState") Then
        Debug.Print LOG_PREFIX & "WARNING: APP_LookupClientState table not found"
        ' Don't fail validation, but warn
    End If
    
    If ValidateSystemTables Then
        Debug.Print LOG_PREFIX & "System tables validated"
    End If
End Function

' ---------------------------------------------------------------------------
' TestCacheSystem - Comprehensive test routine
' ---------------------------------------------------------------------------
Public Sub TestCacheSystem()
    On Error Resume Next
    
    Debug.Print String(70, "=")
    Debug.Print "CACHE SYSTEM TEST"
    Debug.Print String(70, "=")
    
    ' Test 1: System tables
    Debug.Print vbCrLf & "Test 1: System Tables"
    If ValidateSystemTables() Then
        Debug.Print "  ✓ PASS"
    Else
        Debug.Print "  ✗ FAIL"
    End If
    
    ' Test 2: Cache initialization
    Debug.Print vbCrLf & "Test 2: Cache Initialization"
    If Not gLocalTableCache Is Nothing Then
        Debug.Print "  ✓ PASS"
    Else
        Debug.Print "  ✗ FAIL"
    End If
    
    ' Test 3: Configuration count
    Debug.Print vbCrLf & "Test 3: Configuration"
    Dim configCount As Long
    configCount = DCount("*", "APP_LookupConfig", "IsActive=True")
    Debug.Print "  Configured tables: " & configCount
    If configCount > 0 Then
        Debug.Print "  ✓ PASS"
    Else
        Debug.Print "  ✗ FAIL - No tables configured"
    End If
    
    ' Test 4: Show statistics
    Debug.Print vbCrLf & "Test 4: Statistics"
    Debug.Print GetCacheStatistics()
    
    Debug.Print vbCrLf & String(70, "=")
    Debug.Print "TEST COMPLETE"
    Debug.Print String(70, "=")
End Sub

' ============================================================================
' HELPER FUNCTIONS
' ============================================================================

Private Function GetLocalTableName(LookupName As String) As String
    On Error Resume Next
    GetLocalTableName = Nz(DLookup("LocalTableName", "APP_LookupConfig", _
                           "LookupName='" & SQLSafe(LookupName) & "' AND IsActive=True"), "")
End Function

Private Function TableExists(tableName As String) As Boolean
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = CurrentDb.TableDefs(tableName)
    TableExists = (Err.Number = 0)
End Function

Private Function SQLSafe(Value As String) As String
    SQLSafe = Replace(Value, "'", "''")
End Function
