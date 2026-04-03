-- M365 Dashboard Database Schema
-- Run this script against the Azure SQL Database to create the required tables

-- Create UserSettings table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='UserSettings' AND xtype='U')
CREATE TABLE [UserSettings] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [UserId] NVARCHAR(100) NOT NULL,
    [Theme] NVARCHAR(20) NOT NULL DEFAULT 'system',
    [RefreshIntervalSeconds] INT NOT NULL DEFAULT 300,
    [DateRangePreference] NVARCHAR(20) NOT NULL DEFAULT 'last30days',
    [ShowWelcomeMessage] BIT NOT NULL DEFAULT 1,
    [CompactMode] BIT NOT NULL DEFAULT 0,
    [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    [UpdatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    CONSTRAINT [PK_UserSettings] PRIMARY KEY ([Id])
);
GO

-- Create unique index on UserId
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_UserSettings_UserId')
CREATE UNIQUE INDEX [IX_UserSettings_UserId] ON [UserSettings] ([UserId]);
GO

-- Create WidgetConfigurations table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='WidgetConfigurations' AND xtype='U')
CREATE TABLE [WidgetConfigurations] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [UserId] NVARCHAR(100) NOT NULL,
    [WidgetType] NVARCHAR(50) NOT NULL,
    [IsEnabled] BIT NOT NULL DEFAULT 1,
    [DisplayOrder] INT NOT NULL DEFAULT 0,
    [GridColumn] INT NOT NULL DEFAULT 0,
    [GridRow] INT NOT NULL DEFAULT 0,
    [GridWidth] INT NOT NULL DEFAULT 1,
    [GridHeight] INT NOT NULL DEFAULT 1,
    [CustomSettings] NVARCHAR(MAX) NULL,
    [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    [UpdatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    CONSTRAINT [PK_WidgetConfigurations] PRIMARY KEY ([Id])
);
GO

-- Create unique index on UserId + WidgetType
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_WidgetConfigurations_UserId_WidgetType')
CREATE UNIQUE INDEX [IX_WidgetConfigurations_UserId_WidgetType] ON [WidgetConfigurations] ([UserId], [WidgetType]);
GO

-- Create CachedMetrics table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='CachedMetrics' AND xtype='U')
CREATE TABLE [CachedMetrics] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [MetricType] NVARCHAR(100) NOT NULL,
    [TenantId] NVARCHAR(100) NOT NULL,
    [Data] NVARCHAR(MAX) NOT NULL,
    [CachedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    [ExpiresAt] DATETIME2 NOT NULL,
    CONSTRAINT [PK_CachedMetrics] PRIMARY KEY ([Id])
);
GO

-- Create indexes on CachedMetrics
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_CachedMetrics_MetricType_TenantId')
CREATE INDEX [IX_CachedMetrics_MetricType_TenantId] ON [CachedMetrics] ([MetricType], [TenantId]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_CachedMetrics_ExpiresAt')
CREATE INDEX [IX_CachedMetrics_ExpiresAt] ON [CachedMetrics] ([ExpiresAt]);
GO

-- Create DashboardLayouts table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='DashboardLayouts' AND xtype='U')
CREATE TABLE [DashboardLayouts] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [UserId] NVARCHAR(100) NOT NULL,
    [Name] NVARCHAR(100) NOT NULL,
    [IsDefault] BIT NOT NULL DEFAULT 0,
    [LayoutJson] NVARCHAR(MAX) NOT NULL,
    [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    [UpdatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    CONSTRAINT [PK_DashboardLayouts] PRIMARY KEY ([Id])
);
GO

-- Create index on DashboardLayouts
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_DashboardLayouts_UserId')
CREATE INDEX [IX_DashboardLayouts_UserId] ON [DashboardLayouts] ([UserId]);
GO

-- Create AuditLogs table
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='AuditLogs' AND xtype='U')
CREATE TABLE [AuditLogs] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [UserId] NVARCHAR(100) NOT NULL,
    [Action] NVARCHAR(100) NOT NULL,
    [Details] NVARCHAR(MAX) NULL,
    [IpAddress] NVARCHAR(50) NULL,
    [Timestamp] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
    CONSTRAINT [PK_AuditLogs] PRIMARY KEY ([Id])
);
GO

-- Create indexes on AuditLogs
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_AuditLogs_UserId')
CREATE INDEX [IX_AuditLogs_UserId] ON [AuditLogs] ([UserId]);
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_AuditLogs_Timestamp')
CREATE INDEX [IX_AuditLogs_Timestamp] ON [AuditLogs] ([Timestamp]);
GO

-- Create EF Migrations History table (so EF Core thinks migrations have been applied)
IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='__EFMigrationsHistory' AND xtype='U')
CREATE TABLE [__EFMigrationsHistory] (
    [MigrationId] NVARCHAR(150) NOT NULL,
    [ProductVersion] NVARCHAR(32) NOT NULL,
    CONSTRAINT [PK___EFMigrationsHistory] PRIMARY KEY ([MigrationId])
);
GO

-- Insert a migration record
IF NOT EXISTS (SELECT * FROM [__EFMigrationsHistory] WHERE [MigrationId] = '20241224_InitialCreate')
INSERT INTO [__EFMigrationsHistory] ([MigrationId], [ProductVersion])
VALUES ('20241224_InitialCreate', '8.0.0');
GO

PRINT 'Database schema created successfully!';
GO
