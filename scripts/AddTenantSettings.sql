-- Migration: AddTenantSettings
-- Creates the TenantSettings table for storing tenant-level configuration

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'TenantSettings')
BEGIN
    CREATE TABLE [TenantSettings] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [TenantId] NVARCHAR(100) NOT NULL,
        [SettingKey] NVARCHAR(100) NOT NULL,
        [SettingValue] NVARCHAR(MAX) NOT NULL,
        [Description] NVARCHAR(500) NULL,
        [LastModifiedBy] NVARCHAR(100) NULL,
        [CreatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        [UpdatedAt] DATETIME2 NOT NULL DEFAULT GETUTCDATE(),
        CONSTRAINT [PK_TenantSettings] PRIMARY KEY ([Id])
    );

    -- Create unique index on TenantId + SettingKey
    CREATE UNIQUE INDEX [IX_TenantSettings_TenantId_SettingKey] 
        ON [TenantSettings] ([TenantId], [SettingKey]);

    PRINT 'TenantSettings table created successfully.';
END
ELSE
BEGIN
    PRINT 'TenantSettings table already exists.';
END
GO
