-- ============================================================================
-- ENHANCED SCHEMA FOR LOCAL TABLE CACHE SYSTEM
-- ============================================================================

-- Table: APP_LookupConfig (Enhanced)
CREATE TABLE APP_LookupConfig (
    LookupName              TEXT(50) PRIMARY KEY,
    Description             TEXT(255),
    AccessQuery             MEMO,           -- Query for Access backend
    SQLServerQuery          MEMO,           -- Query for SQL Server backend
    LocalTableName          TEXT(50),       -- Name of local cache table
    IsActive                BIT DEFAULT 1,
    IsStatic                BIT DEFAULT 1,  -- Static tables are cached locally
    
    -- REFRESH CONTROL
    RefreshOnStartup        BIT DEFAULT 0,  -- Load immediately on app start
    AllowManualRefresh      BIT DEFAULT 1,  -- Allow user-triggered refresh
    RefreshIntervalHours    INTEGER DEFAULT 0,  -- 0 = no time-based refresh
    RefreshVersion          LONG DEFAULT 1, -- Increment to force refresh
    
    -- LOAD BEHAVIOR
    LoadOrder               INTEGER DEFAULT 100,    -- Lower = loads first
    LazyLoad                BIT DEFAULT 1,  -- Wait until first use
    TruncateBeforeLoad      BIT DEFAULT 1,  -- Clear before loading
    UseTransaction          BIT DEFAULT 1,  -- Wrap in transaction
    DeleteOnShutdown        BIT DEFAULT 0,  -- Remove on app close
    
    -- METADATA
    CreatedOn               DATETIME DEFAULT Now(),
    CreatedBy               LONG,
    ModifiedOn              DATETIME,
    ModifiedBy              LONG
);

-- Table: APP_LookupClientState (Enhanced)
CREATE TABLE APP_LookupClientState (
    LookupName              TEXT(50),
    ClientMachineName       TEXT(50),
    LastRefreshVersion      LONG DEFAULT 0,
    LastRefreshDate         DATETIME,
    LastRecordCount         INTEGER DEFAULT 0,
    LastRefreshStatus       TEXT(20) DEFAULT 'Pending',  -- Pending/Success/Failed
    LastRefreshDuration     LONG DEFAULT 0,  -- Milliseconds
    LastErrorMessage        MEMO,
    
    -- Composite Primary Key
    CONSTRAINT PK_ClientState PRIMARY KEY (LookupName, ClientMachineName)
);

-- Index for performance
CREATE INDEX IDX_LookupConfig_Active ON APP_LookupConfig(IsActive, IsStatic);
CREATE INDEX IDX_ClientState_Lookup ON APP_LookupClientState(LookupName);

-- ============================================================================
-- SAMPLE CONFIGURATION DATA
-- ============================================================================

INSERT INTO APP_LookupConfig (
    LookupName, 
    Description, 
    AccessQuery, 
    LocalTableName,
    RefreshOnStartup,
    LazyLoad,
    RefreshIntervalHours,
    DeleteOnShutdown
) VALUES (
    'Customers',
    'Customer master list',
    'SELECT CustomerID, CustomerName FROM tblCustomers WHERE IsActive=True',
    'Local_Customers',
    0,      -- Don't load on startup
    1,      -- Load when first used
    24,     -- Refresh every 24 hours
    0       -- Keep after shutdown
);

INSERT INTO APP_LookupConfig (
    LookupName, 
    Description, 
    AccessQuery, 
    LocalTableName,
    RefreshOnStartup,
    LazyLoad,
    RefreshIntervalHours,
    DeleteOnShutdown
) VALUES (
    'TempData',
    'Temporary session data',
    'SELECT * FROM tblTempData',
    'Local_TempData',
    1,      -- Load on startup
    0,      -- Not lazy (pre-load)
    0,      -- No time-based refresh
    1       -- Delete on shutdown
);
