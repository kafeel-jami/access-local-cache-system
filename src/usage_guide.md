# Enhanced Local Table Cache System - Usage Guide

## Overview

This enhanced system provides **lazy-loading**, **version-controlled caching** with proper respect for all configuration flags.

## Key Improvements

### 1. **Lazy Loading** ✅
- Tables are **NOT** created until first use
- Only tables marked `RefreshOnStartup=True` load at startup
- Dramatically reduces startup time

### 2. **Proper Flag Handling** ✅
- `RefreshOnStartup`: Only loads these tables on app start
- `LazyLoad`: Creates/loads table on first access
- `DeleteOnShutdown`: Actually deletes only marked tables
- `RefreshIntervalHours`: Properly checks time-based refresh
- `RefreshVersion`: Triggers refresh when version incremented

### 3. **Performance Optimizations** ✅
- Connection pooling with retry logic
- Memory-based filtering in FindAsYouTypeCombo
- Transaction support for data loading
- Proper error handling throughout

## Configuration Examples

### Example 1: Core Reference Data (Keep Always)
```sql
INSERT INTO APP_LookupConfig (
    LookupName, LocalTableName, AccessQuery,
    RefreshOnStartup, LazyLoad, RefreshIntervalHours, DeleteOnShutdown
) VALUES (
    'Countries',
    'Local_Countries', 
    'SELECT CountryID, CountryName FROM tblCountries',
    0,    -- Don't load at startup
    1,    -- Load when first used
    0,    -- No time-based refresh (stable data)
    0     -- Keep table on shutdown
);
```

### Example 2: Daily Updated Data
```sql
INSERT INTO APP_LookupConfig (
    LookupName, LocalTableName, AccessQuery,
    RefreshOnStartup, LazyLoad, RefreshIntervalHours, DeleteOnShutdown
) VALUES (
    'DailyRates',
    'Local_DailyRates',
    'SELECT * FROM tblExchangeRates WHERE IsActive=True',
    1,    -- Load at startup (always need fresh)
    0,    -- Not lazy (pre-load)
    24,   -- Refresh every 24 hours
    0     -- Keep table
);
```

### Example 3: Session-Only Data
```sql
INSERT INTO APP_LookupConfig (
    LookupName, LocalTableName, AccessQuery,
    RefreshOnStartup, LazyLoad, RefreshIntervalHours, DeleteOnShutdown
) VALUES (
    'SessionData',
    'Local_SessionData',
    'SELECT * FROM tblTempData WHERE UserID=' & CurrentUserID,
    0,    -- Don't load at startup
    1,    -- Load when used
    0,    -- No auto-refresh
    1     -- DELETE on shutdown (temporary)
);
```

## Usage in Forms

### Binding a Combo Box (Lazy Loading)

```vba
Option Compare Database
Option Explicit

Private faytCustomers As FindAsYouTypeCombo

Private Sub Form_Open(Cancel As Integer)
    ' This will automatically load the table if not already cached
    Set faytCustomers = BindComboFromCache(Me.cboCustomer, "Customers")
End Sub
```

### Manual Refresh from Button

```vba
Private Sub btnRefresh_Click()
    If ForceRefresh("Customers") Then
        MsgBox "Customers refreshed!", vbInformation
        ' Requery combo if needed
        Me.cboCustomer.Requery
    End If
End Sub
```

### Checking Table Status

```vba
Private Sub btnCheckStatus_Click()
    Debug.Print GetTableInfo("Customers")
    ' Or show in message box:
    MsgBox GetTableInfo("Customers"), vbInformation
End Sub
```

## Startup Code (in AutoExec or Form_Open)

```vba
Private Sub Form_Open(Cancel As Integer)
    ' This is all you need!
    Call ApplicationStartup
    
    ' Tables marked RefreshOnStartup=True will load
    ' All others wait until first use (lazy loading)
End Sub
```

## Shutdown Code (in Form_Close or AutoExec)

```vba
Private Sub Form_Close()
    ' This deletes ONLY tables marked DeleteOnShutdown=True
    Call ApplicationShutdown
End Sub
```

## Forcing Version Refresh

When you need to force all clients to refresh:

```sql
-- Increment RefreshVersion to trigger refresh
UPDATE APP_LookupConfig 
SET RefreshVersion = RefreshVersion + 1,
    ModifiedOn = Now()
WHERE LookupName = 'Customers';
```

Next time any user accesses "Customers", it will automatically refresh.

## Monitoring & Diagnostics

### View Cache Statistics
```vba
Debug.Print GetCacheStatistics()
```

### Full Cache Report
```vba
Call ShowCacheReport
```

### Test System
```vba
Call TestCacheSystem
```

### View Specific Table Info
```vba
Debug.Print GetTableInfo("Customers")
```

## Architecture Flow

```
Application Startup
    ↓
Initialize Connection
    ↓
Initialize Cache Object
    ↓
Load ONLY RefreshOnStartup=True tables
    ↓
[User opens form with combo]
    ↓
BindComboFromCache called
    ↓
EnsureTableReady checks:
  - Does table exist? → Create if needed
  - Is refresh needed? → Check version & time
  - Load/refresh data if needed
    ↓
Combo bound and ready
```

## Performance Tips

1. **Set LazyLoad=True** for rarely-used tables
2. **Set RefreshOnStartup=False** unless critical
3. **Use RefreshIntervalHours** only for volatile data
4. **Set DeleteOnShutdown=True** for temporary data
5. **Increment RefreshVersion** strategically (not every change)

## Troubleshooting

### Table Not Loading
1. Check `APP_LookupConfig`: `IsActive=True`?
2. Check debug output for errors
3. Run `TestCacheSystem` to diagnose

### Slow Startup
1. Too many `RefreshOnStartup=True` tables
2. Change to `LazyLoad=True` where possible

### Data Not Refreshing
1. Check `RefreshVersion` - increment to force
2. Check `RefreshIntervalHours` setting
3. Use `ForceRefresh(LookupName)` manually

## Best Practices

1. ✅ Use **lazy loading** for most tables
2. ✅ Reserve **RefreshOnStartup** for critical data only
3. ✅ Set **DeleteOnShutdown** for temporary/user-specific data
4. ✅ Use **RefreshVersion** for controlled updates
5. ✅ Set **RefreshIntervalHours** for time-sensitive data
6. ✅ Always call `EnsureTableReady()` before direct table access

## Migration from Old System

If migrating from the old system:

1. **Update Schema**: Add new columns (`LazyLoad`)
2. **Set Defaults**: `LazyLoad=1`, `RefreshOnStartup=0` for most tables
3. **Identify Critical Tables**: Set `RefreshOnStartup=1` only where needed
4. **Mark Temporary Tables**: Set `DeleteOnShutdown=1` appropriately
5. **Update Startup Code**: Use new `ApplicationStartup()`
6. **Test**: Run `TestCacheSystem` after migration

## Summary

The enhanced system provides:
- ⚡ **Faster startup** (lazy loading)
- 🎯 **Proper flag handling** (DeleteOnShutdown, RefreshOnStartup)
- 🔄 **Smart refresh logic** (version + time-based)
- 📊 **Better monitoring** (detailed state tracking)
- 🛡️ **Robust error handling** (retry logic, transactions)
- 🚀 **Performance optimized** (connection pooling, caching)
