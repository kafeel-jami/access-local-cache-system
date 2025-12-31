# Access Local Cache System

High-performance local caching system for Microsoft Access applications with SQL Server migration support.

## Features

- Lazy loading of lookup tables
- Backend-agnostic (Access/SQL Server)
- Find-as-you-type combo boxes
- Configurable refresh strategies
- Performance optimization

## Installation

1. Create the schema tables in your Access database
2. Import the VBA modules and class
3. Configure your lookup tables
4. Initialize the system in application startup

## Usage

```vba
' Initialize cache system
ApplicationStartup

' Ensure table is ready (loads if needed)
EnsureTableReady "Customers"

' Bind combo box to cached table
Dim combo As FindAsYouTypeCombo
Set combo = BindComboFromCache(Me.MyCombo, "Customers", "CustomerName")