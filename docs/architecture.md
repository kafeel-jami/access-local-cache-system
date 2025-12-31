# System Architecture

## Overview

The Access Local Cache System is designed as a multi-layered architecture that provides high-performance caching for Microsoft Access applications while supporting seamless migration from Access to SQL Server backends.

## Component Architecture

### 1. Database Schema Layer
The foundation consists of two configuration tables that control caching behavior:

**APP_LookupConfig Table**
- Stores cache configuration for each lookup table
- Supports dual queries for Access and SQL Server backends
- Controls refresh strategies and load behavior
- Manages per-table settings (startup, lazy loading, intervals)

**APP_LookupClientState Table**
- Tracks refresh status per client machine
- Records timing, counts, and error information
- Supports distributed environments with machine-specific state

### 2. Connection Management Layer (modConnection)
- **Backend Abstraction**: Supports both Access (ACE/JET) and SQL Server
- **Connection Pooling**: Optimized connection management with retry logic
- **Disconnected Operations**: Supports disconnected recordsets for caching
- **Transaction Support**: Built-in transaction management

### 3. Cache Engine Layer (clsLocalTableCache)
- **Lazy Loading**: Tables loaded only when first accessed
- **Refresh Logic**: Multiple strategies (time-based, version-based, manual)
- **State Management**: Per-machine refresh tracking
- **Error Recovery**: Comprehensive error handling and recovery

### 4. API Layer (modLocalTableManager)
- **High-Level Interface**: Simple functions for cache operations
- **Lifecycle Management**: Startup and shutdown handling
- **UI Integration**: Combo box binding with find-as-you-type
- **Monitoring**: Built-in diagnostics and reporting

## Data Flow
Application Request
       ↓
EnsureTableReady()
       ↓
Check Local Table Exists
       ↓
Is Refresh Needed? → Yes → Load from Source → Update Local Table
       ↓ No
Use Cached Data
       ↓
Return to Application


## Design Principles

### Performance Optimization
- Memory caching of frequently accessed data
- Disconnected recordsets to reduce database connections
- Lazy loading to minimize startup impact
- Efficient query execution strategies

### Scalability
- Machine-specific state management
- Configurable refresh intervals
- Connection pooling and optimization
- Resource cleanup and management

### Maintainability
- Database-driven configuration
- Clear separation of concerns
- Comprehensive error handling
- Built-in monitoring and diagnostics

## Migration Support

The architecture is specifically designed to support backend migration:
- Dual query support in configuration tables
- Backend-agnostic connection management
- Transparent switching between Access and SQL Server
- Consistent performance during migration