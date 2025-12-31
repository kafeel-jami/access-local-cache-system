# Installation Guide

## Prerequisites

- Microsoft Access 2010 or later
- Microsoft ADO libraries (included with Access)
- VBA development environment
- Appropriate database permissions

## Step-by-Step Installation

### 1. Create Database Schema

Execute the following SQL in your Access database:

```sql
-- Table: APP_LookupConfig
CREATE TABLE APP_LookupConfig (
    LookupName TEXT(50) PRIMARY KEY,
    Description TEXT(255),
    AccessQuery MEMO,
    SQLServerQuery MEMO,
    LocalTableName TEXT(50),
    IsActive BIT DEFAULT 1,
    IsStatic BIT DEFAULT 1,
    RefreshOnStartup BIT DEFAULT 0,
    AllowManualRefresh BIT DEFAULT 1,
    RefreshIntervalHours INTEGER DEFAULT 0,
    RefreshVersion LONG DEFAULT 1,
    LoadOrder INTEGER DEFAULT 100,
    LazyLoad BIT DEFAULT 1,
    TruncateBeforeLoad BIT DEFAULT 1,
    UseTransaction BIT DEFAULT 1,
    DeleteOnShutdown BIT DEFAULT 0,
    CreatedOn DATETIME DEFAULT Now(),
    CreatedBy LONG,
    ModifiedOn DATETIME,
    ModifiedBy LONG
);

-- Table: APP_LookupClientState
CREATE TABLE APP_LookupClientState (
    LookupName TEXT(50),
    ClientMachineName TEXT(50),
    LastRefreshVersion LONG DEFAULT 0,
    LastRefreshDate DATETIME,
    LastRecordCount INTEGER DEFAULT 0,
    LastRefreshStatus TEXT(20) DEFAULT 'Pending',
    LastRefreshDuration LONG DEFAULT 0,
    LastErrorMessage MEMO,
    CONSTRAINT PK_ClientState PRIMARY KEY (LookupName, ClientMachineName)
);

-- Create indexes
CREATE INDEX IDX_LookupConfig_Active ON APP_LookupConfig(IsActive, IsStatic);
CREATE INDEX IDX_ClientState_Lookup ON APP_LookupClientState(LookupName);