
## **docs/migration-guide.md**

```markdown
# Migration Guide: Access to SQL Server

## Overview

This guide provides step-by-step instructions for migrating your Access application from Access (ACE/JET) backend to SQL Server while maintaining the caching system functionality.

## Prerequisites

- Access to SQL Server instance
- Database backup of your Access backend
- Test environment for migration validation
- Understanding of your current Access queries

## Migration Strategy

The caching system supports a phased migration approach:

### Phase 1: Preparation
- Configure dual queries in APP_LookupConfig
- Test SQL Server connections
- Validate query compatibility

### Phase 2: Parallel Operation
- Run both backends simultaneously
- Validate data consistency
- Monitor performance

### Phase 3: Backend Switch
- Switch global backend configuration
- Monitor application performance
- Clean up Access backend (optional)

## Step-by-Step Migration

### 1. Database Preparation

**SQL Server Setup:**
1. Create SQL Server database
2. Import Access tables using SQL Server Import/Export Wizard
3. Create indexes matching Access database
4. Test basic connectivity

**Access Database Updates:**
1. Add SQL Server queries to APP_LookupConfig
2. Test connection strings in modConnection
3. Update gBackendType as needed

### 2. Connection Configuration

Update connection strings in `modConnection`:

```vba
Private Function GetConnectionString() As String
    Select Case gBackendType
        Case Backend_SQLServer
            GetConnectionString = _
                "Provider=SQLOLEDB;" & _
                "Data Source=SERVER_NAME;" & _
                "Initial Catalog=DATABASE_NAME;" & _
                "Integrated Security=SSPI;" & _
                "Connection Timeout=30;"
                
        Case Backend_Access
            GetConnectionString = _
                "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Data Source=" & CurrentProject.Path & "\YourAccessDB.accdb;" & _
                "Mode=Share Deny None;"
    End Select
End Function