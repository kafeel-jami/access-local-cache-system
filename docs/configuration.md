
## **docs/configuration.md**

```markdown
# Configuration Guide

## Overview

The caching system is controlled through the `APP_LookupConfig` table, which provides flexible configuration options for each cached lookup table.

## Configuration Fields

### Basic Information
- **LookupName**: Unique identifier for the cache configuration
- **Description**: Human-readable description of the lookup
- **LocalTableName**: Name of the local cache table to create

### Query Configuration
- **AccessQuery**: SQL query for Access backend
- **SQLServerQuery**: SQL query for SQL Server backend

### Refresh Control
- **RefreshOnStartup**: Load table immediately on application start (0/1)
- **AllowManualRefresh**: Allow user-triggered refresh (0/1)
- **RefreshIntervalHours**: Hours between automatic refresh (0 = no refresh)
- **RefreshVersion**: Version number to force refresh (increment to refresh)

### Load Behavior
- **LoadOrder**: Order to load tables (lower numbers load first)
- **LazyLoad**: Wait until first use before loading (0/1)
- **TruncateBeforeLoad**: Clear before loading new data (0/1)
- **UseTransaction**: Wrap load in transaction (0/1)
- **DeleteOnShutdown**: Remove table on application close (0/1)

### Status
- **IsActive**: Enable/disable this configuration (0/1)
- **IsStatic**: Mark as static lookup table (0/1)

## Configuration Examples

### Basic Lookup Table
```sql
INSERT INTO APP_LookupConfig (
    LookupName, Description, AccessQuery, LocalTableName,
    RefreshOnStartup, RefreshIntervalHours, LoadOrder, LazyLoad
) VALUES (
    'Customers', 'Customer master list',
    'SELECT CustomerID, CustomerName FROM tblCustomers WHERE IsActive=True',
    'Local_Customers', 0, 24, 10, 1
);