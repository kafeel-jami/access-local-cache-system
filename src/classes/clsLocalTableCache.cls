' ============================================================================
' CLASS: clsLocalTableCache (Enhanced v2.0)
' PURPOSE: Core cache engine with lazy loading and version control
' STRATEGY: Load on demand, respect configuration flags
' ============================================================================
Option Compare Database
Option Explicit

' === Private State ===
Private mMachineName As String
Private mIsInitialized As Boolean

' ============================================================================
' INITIALIZATION
' ============================================================================

Private Sub Class_Initialize()
    mMachineName = Environ("COMPUTERNAME")
    mIsInitialized = True
End Sub

' ============================================================================
' PUBLIC PROPERTIES
' ============================================================================

Public Property Get IsInitialized() As Boolean
    IsInitialized = mIsInitialized
End Property

Public Property Get MachineName() As String
    MachineName = mMachineName
End Property

' ============================================================================
' MAIN ENTRY POINTS
' ============================================================================

' -----------------------------------------------------------------------
' LoadStartupTables - Load only tables marked for startup
' -----------------------------------------------------------------------
Public Sub LoadStartupTables()
    On Error GoTo ErrHandler
    
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim loadedCount As Long
    Dim failedCount As Long
    
    sql = "SELECT LookupName FROM APP_LookupConfig " & _
          "WHERE IsActive=True AND IsStatic=True AND RefreshOnStartup=True " & _
          "ORDER BY LoadOrder"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.EOF And rs.BOF Then
        Debug.Print "Cache> No startup tables configured"
        rs.Close
        Exit Sub
    End If
    
    Debug.Print "Cache> Loading startup tables..."
    
    Do Until rs.EOF
        If EnsureTableReady(rs!LookupName) Then
            loadedCount = loadedCount + 1
            Debug.Print "Cache>   ✓ " & rs!LookupName
        Else
            failedCount = failedCount + 1
            Debug.Print "Cache>   ✗ " & rs!LookupName & " (failed)"
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    Debug.Print "Cache> Startup complete: " & loadedCount & " loaded, " & failedCount & " failed"
    
    Exit Sub
    
ErrHandler:
    If Not rs Is Nothing Then rs.Close
    Debug.Print "Cache> ERROR in LoadStartupTables: " & Err.Description
End Sub

' -----------------------------------------------------------------------
' EnsureTableReady - Main function: ensures table exists and is current
' This is called whenever a table is accessed
' -----------------------------------------------------------------------
Public Function EnsureTableReady(LookupName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim cfg As LookupConfig
    
    ' Get configuration
    If Not GetConfig(LookupName, cfg) Then
        Debug.Print "Cache> Config not found: " & LookupName
        Exit Function
    End If
    
    ' Step 1: Ensure table structure exists
    If Not TableExists(cfg.LocalTableName) Then
        If Not CreateLocalTable(cfg) Then
            Debug.Print "Cache> Failed to create table: " & cfg.LocalTableName
            Exit Function
        End If
    End If
    
    ' Step 2: Check if refresh needed
    If IsRefreshNeeded(cfg) Then
        If Not RefreshTableData(cfg) Then
            Debug.Print "Cache> Failed to refresh: " & LookupName
            Exit Function
        End If
    End If
    
    EnsureTableReady = True
    Exit Function
    
ErrHandler:
    Debug.Print "Cache> ERROR in EnsureTableReady(" & LookupName & "): " & Err.Description
    EnsureTableReady = False
End Function

' -----------------------------------------------------------------------
' ForceRefresh - Manually refresh a specific table
' -----------------------------------------------------------------------
Public Function ForceRefresh(LookupName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim cfg As LookupConfig
    
    If Not GetConfig(LookupName, cfg) Then
        Exit Function
    End If
    
    ' Ensure table exists first
    If Not TableExists(cfg.LocalTableName) Then
        If Not CreateLocalTable(cfg) Then
            Exit Function
        End If
    End If
    
    ' Force refresh
    ForceRefresh = RefreshTableData(cfg)
    
    Exit Function
    
ErrHandler:
    Debug.Print "Cache> ERROR in ForceRefresh: " & Err.Description
    ForceRefresh = False
End Function

' ============================================================================
' REFRESH LOGIC
' ============================================================================

' -----------------------------------------------------------------------
' IsRefreshNeeded - Determines if table needs refreshing
' -----------------------------------------------------------------------
Private Function IsRefreshNeeded(cfg As LookupConfig) As Boolean
    On Error Resume Next
    
    Dim clientState As ClientState
    
    ' Get client state
    If Not GetClientState(cfg.LookupName, clientState) Then
        ' No state = needs refresh
        IsRefreshNeeded = True
        Exit Function
    End If
    
    ' Check 1: Version mismatch
    If cfg.RefreshVersion > clientState.LastRefreshVersion Then
        Debug.Print "Cache> Version refresh needed: " & cfg.LookupName & _
                    " (Config: " & cfg.RefreshVersion & ", Client: " & clientState.LastRefreshVersion & ")"
        IsRefreshNeeded = True
        Exit Function
    End If
    
    ' Check 2: Time-based refresh
    If cfg.RefreshIntervalHours > 0 Then
        If IsNull(clientState.LastRefreshDate) Then
            IsRefreshNeeded = True
            Exit Function
        End If
        
        Dim hoursSinceRefresh As Long
        hoursSinceRefresh = DateDiff("h", clientState.LastRefreshDate, Now())
        
        If hoursSinceRefresh >= cfg.RefreshIntervalHours Then
            Debug.Print "Cache> Time-based refresh needed: " & cfg.LookupName & _
                        " (" & hoursSinceRefresh & " hours elapsed)"
            IsRefreshNeeded = True
            Exit Function
        End If
    End If
    
    ' Check 3: Failed last refresh
    If clientState.LastRefreshStatus = "Failed" Then
        Debug.Print "Cache> Retry after failure: " & cfg.LookupName
        IsRefreshNeeded = True
        Exit Function
    End If
    
    IsRefreshNeeded = False
End Function

' -----------------------------------------------------------------------
' RefreshTableData - Loads data from source into local table
' -----------------------------------------------------------------------
Private Function RefreshTableData(cfg As LookupConfig) As Boolean
    On Error GoTo ErrHandler
    
    Dim startTime As Double
    Dim recordsLoaded As Long
    Dim sourceQuery As String
    
    startTime = Timer
    
    ' Get appropriate query
    sourceQuery = GetSourceQuery(cfg)
    If sourceQuery = "" Then
        UpdateClientState cfg.LookupName, False, 0, 0, "No source query defined"
        Exit Function
    End If
    
    Debug.Print "Cache> Refreshing: " & cfg.LookupName
    
    ' Use transaction if configured
    If cfg.UseTransaction Then
        CurrentDb.Execute "BEGIN TRANSACTION", dbFailOnError
    End If
    
    On Error GoTo TransactionError
    
    ' Truncate if configured
    If cfg.TruncateBeforeLoad Then
        CurrentDb.Execute "DELETE FROM [" & cfg.LocalTableName & "]", dbFailOnError
    End If
    
    ' Load data
    CurrentDb.Execute "INSERT INTO [" & cfg.LocalTableName & "] " & sourceQuery, dbFailOnError
    
    ' Count records
    recordsLoaded = DCount("*", cfg.LocalTableName)
    
    ' Commit transaction
    If cfg.UseTransaction Then
        CurrentDb.Execute "COMMIT TRANSACTION", dbFailOnError
    End If
    
    ' Update state
    Dim duration As Long
    duration = CLng((Timer - startTime) * 1000) ' milliseconds
    
    UpdateClientState cfg.LookupName, True, cfg.RefreshVersion, recordsLoaded, "", duration
    
    Debug.Print "Cache>   ✓ Loaded " & recordsLoaded & " records in " & duration & "ms"
    
    RefreshTableData = True
    Exit Function
    
TransactionError:
    If cfg.UseTransaction Then
        On Error Resume Next
        CurrentDb.Execute "ROLLBACK TRANSACTION", dbFailOnError
    End If
    
ErrHandler:
    UpdateClientState cfg.LookupName, False, 0, 0, Err.Description
    Debug.Print "Cache> ERROR refreshing " & cfg.LookupName & ": " & Err.Description
    RefreshTableData = False
End Function

' ============================================================================
' TABLE CREATION
' ============================================================================

' -----------------------------------------------------------------------
' CreateLocalTable - Creates empty local table with proper structure
' -----------------------------------------------------------------------
Private Function CreateLocalTable(cfg As LookupConfig) As Boolean
    On Error GoTo ErrHandler
    
    Dim schemaQuery As String
    
    schemaQuery = GetSourceQuery(cfg)
    If schemaQuery = "" Then
        Err.Raise vbObjectError + 1001, , "No schema query for " & cfg.LookupName
    End If
    
    ' Remove trailing semicolon
    If Right(schemaQuery, 1) = ";" Then
        schemaQuery = Left(schemaQuery, Len(schemaQuery) - 1)
    End If
    
    ' Create empty table with schema
    Debug.Print "Cache> Creating table: " & cfg.LocalTableName
    
    CurrentDb.Execute _
        "SELECT * INTO [" & cfg.LocalTableName & "] " & _
        "FROM (" & schemaQuery & ") AS SourceData WHERE 1=0", _
        dbFailOnError
    
    CreateLocalTable = True
    Exit Function
    
ErrHandler:
    Debug.Print "Cache> ERROR creating table " & cfg.LocalTableName & ": " & Err.Description
    CreateLocalTable = False
End Function

' ============================================================================
' STATE MANAGEMENT
' ============================================================================

' -----------------------------------------------------------------------
' UpdateClientState - Records refresh results
' -----------------------------------------------------------------------
Private Sub UpdateClientState( _
    LookupName As String, _
    Success As Boolean, _
    RefreshVersion As Long, _
    RecordCount As Long, _
    ErrorMsg As String, _
    Optional Duration As Long = 0)
    
    On Error GoTo ErrHandler
    
    Dim sql As String
    
    ' Delete existing state
    sql = "DELETE FROM APP_LookupClientState " & _
          "WHERE LookupName='" & SQLSafe(LookupName) & "' " & _
          "AND ClientMachineName='" & SQLSafe(mMachineName) & "'"
    CurrentDb.Execute sql, dbFailOnError
    
    ' Insert new state
    sql = "INSERT INTO APP_LookupClientState " & _
          "(LookupName, ClientMachineName, LastRefreshVersion, LastRefreshDate, " & _
          "LastRecordCount, LastRefreshStatus, LastRefreshDuration, LastErrorMessage) " & _
          "VALUES (" & _
          "'" & SQLSafe(LookupName) & "', " & _
          "'" & SQLSafe(mMachineName) & "', " & _
          RefreshVersion & ", " & _
          IIf(Success, "Now()", "NULL") & ", " & _
          RecordCount & ", " & _
          "'" & IIf(Success, "Success", "Failed") & "', " & _
          Duration & ", " & _
          "'" & SQLSafe(ErrorMsg) & "')"
    
    CurrentDb.Execute sql, dbFailOnError
    
    Exit Sub
    
ErrHandler:
    Debug.Print "Cache> ERROR updating client state: " & Err.Description
End Sub

' -----------------------------------------------------------------------
' GetClientState - Retrieves current state for this machine
' -----------------------------------------------------------------------
Private Function GetClientState(LookupName As String, ByRef state As ClientState) As Boolean
    On Error Resume Next
    
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM APP_LookupClientState " & _
          "WHERE LookupName='" & SQLSafe(LookupName) & "' " & _
          "AND ClientMachineName='" & SQLSafe(mMachineName) & "'"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.EOF And rs.BOF Then
        rs.Close
        GetClientState = False
        Exit Function
    End If
    
    With state
        .LookupName = Nz(rs!LookupName, "")
        .LastRefreshVersion = Nz(rs!LastRefreshVersion, 0)
        .LastRefreshDate = Nz(rs!LastRefreshDate, #1/1/1900#)
        .LastRecordCount = Nz(rs!LastRecordCount, 0)
        .LastRefreshStatus = Nz(rs!LastRefreshStatus, "")
        .LastErrorMessage = Nz(rs!LastErrorMessage, "")
    End With
    
    rs.Close
    GetClientState = True
End Function

' ============================================================================
' CONFIGURATION HELPERS
' ============================================================================

' -----------------------------------------------------------------------
' GetConfig - Loads configuration for a lookup
' -----------------------------------------------------------------------
Private Function GetConfig(LookupName As String, ByRef cfg As LookupConfig) As Boolean
    On Error Resume Next
    
    Dim rs As DAO.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM APP_LookupConfig " & _
          "WHERE LookupName='" & SQLSafe(LookupName) & "' AND IsActive=True"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.EOF And rs.BOF Then
        rs.Close
        GetConfig = False
        Exit Function
    End If
    
    With cfg
        .LookupName = Nz(rs!LookupName, "")
        .LocalTableName = Nz(rs!LocalTableName, "")
        .AccessQuery = Nz(rs!AccessQuery, "")
        .SQLServerQuery = Nz(rs!SQLServerQuery, "")
        .RefreshOnStartup = Nz(rs!RefreshOnStartup, False)
        .RefreshIntervalHours = Nz(rs!RefreshIntervalHours, 0)
        .RefreshVersion = Nz(rs!RefreshVersion, 1)
        .TruncateBeforeLoad = Nz(rs!TruncateBeforeLoad, True)
        .UseTransaction = Nz(rs!UseTransaction, True)
        .DeleteOnShutdown = Nz(rs!DeleteOnShutdown, False)
        .LazyLoad = Nz(rs!LazyLoad, True)
    End With
    
    rs.Close
    GetConfig = True
End Function

' -----------------------------------------------------------------------
' GetSourceQuery - Returns appropriate query for current backend
' -----------------------------------------------------------------------
Private Function GetSourceQuery(cfg As LookupConfig) As String
    If gBackendType = Backend_SQLServer Then
        GetSourceQuery = cfg.SQLServerQuery
    Else
        GetSourceQuery = cfg.AccessQuery
    End If
End Function

' ============================================================================
' SHUTDOWN CLEANUP
' ============================================================================

' -----------------------------------------------------------------------
' CleanupOnShutdown - Deletes tables marked for deletion
' -----------------------------------------------------------------------
Public Function CleanupOnShutdown() As Long
    On Error Resume Next
    
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim deleteCount As Long
    
    sql = "SELECT LocalTableName FROM APP_LookupConfig " & _
          "WHERE IsActive=True AND IsStatic=True AND DeleteOnShutdown=True"
    
    Set rs = CurrentDb.OpenRecordset(sql)
    
    If rs.EOF And rs.BOF Then
        rs.Close
        CleanupOnShutdown = 0
        Exit Function
    End If
    
    Debug.Print "Cache> Cleaning up shutdown tables..."
    
    Do Until rs.EOF
        If TableExists(rs!LocalTableName) Then
            CurrentDb.TableDefs.Delete rs!LocalTableName
            If Err.Number = 0 Then
                deleteCount = deleteCount + 1
                Debug.Print "Cache>   ✓ Deleted: " & rs!LocalTableName
            Else
                Debug.Print "Cache>   ✗ Failed to delete: " & rs!LocalTableName
                Err.Clear
            End If
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    CleanupOnShutdown = deleteCount
End Function

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

Private Function TableExists(tableName As String) As Boolean
    On Error Resume Next
    Dim tdf As DAO.TableDef
    Set tdf = CurrentDb.TableDefs(tableName)
    TableExists = (Err.Number = 0)
End Function

Private Function SQLSafe(Value As String) As String
    SQLSafe = Replace(Value, "'", "''")
End Function

' ============================================================================
' TYPE DEFINITIONS
' ============================================================================

Private Type LookupConfig
    LookupName As String
    LocalTableName As String
    AccessQuery As String
    SQLServerQuery As String
    RefreshOnStartup As Boolean
    RefreshIntervalHours As Long
    RefreshVersion As Long
    TruncateBeforeLoad As Boolean
    UseTransaction As Boolean
    DeleteOnShutdown As Boolean
    LazyLoad As Boolean
End Type

Private Type ClientState
    LookupName As String
    LastRefreshVersion As Long
    LastRefreshDate As Date
    LastRecordCount As Long
    LastRefreshStatus As String
    LastErrorMessage As String
End Type
