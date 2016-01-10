
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 01/05/2016 19:27:42
-- Generated from EDMX file: C:\GitHub\BertPnPSitesCore\Core\Tools\OfficeDevPnP.Core.Tools.UnitTest\OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions\SQL\TestModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [PnP];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FK_TestResultTestResultMessage]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[TestResultMessageSet] DROP CONSTRAINT [FK_TestResultTestResultMessage];
GO
IF OBJECT_ID(N'[dbo].[FK_TestRunTestResult]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[TestResultSet] DROP CONSTRAINT [FK_TestRunTestResult];
GO
IF OBJECT_ID(N'[dbo].[FK_TestConfigurationTestRun]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[TestRunSet] DROP CONSTRAINT [FK_TestConfigurationTestRun];
GO
IF OBJECT_ID(N'[dbo].[FK_TestConfigurationTestAuthentication]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[TestConfigurationSet] DROP CONSTRAINT [FK_TestConfigurationTestAuthentication];
GO
IF OBJECT_ID(N'[dbo].[FK_TestConfigurationTestConfigurationProperty]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[TestConfigurationPropertySet] DROP CONSTRAINT [FK_TestConfigurationTestConfigurationProperty];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[TestRunSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TestRunSet];
GO
IF OBJECT_ID(N'[dbo].[TestResultSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TestResultSet];
GO
IF OBJECT_ID(N'[dbo].[TestResultMessageSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TestResultMessageSet];
GO
IF OBJECT_ID(N'[dbo].[TestConfigurationSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TestConfigurationSet];
GO
IF OBJECT_ID(N'[dbo].[TestAuthenticationSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TestAuthenticationSet];
GO
IF OBJECT_ID(N'[dbo].[TestConfigurationPropertySet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TestConfigurationPropertySet];
GO
IF OBJECT_ID(N'[dbo].[FileTrackingSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[FileTrackingSet];
GO
IF OBJECT_ID(N'[dbo].[FileTrackingBaselineSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[FileTrackingBaselineSet];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'TestRunSet'
CREATE TABLE [dbo].[TestRunSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [TestConfigurationId] int  NOT NULL,
    [TestDate] datetime  NOT NULL,
    [TestTime] time  NULL,
    [Build] nvarchar(max)  NOT NULL,
    [Status] int  NOT NULL,
    [TestWasAborted] bit  NOT NULL,
    [TestWasCancelled] bit  NOT NULL,
    [TestsPassed] int  NULL,
    [TestsSkipped] int  NULL,
    [TestsFailed] int  NULL,
    [TestsNotFound] int  NULL,
    [MSBuildLog] nvarchar(max)  NULL
);
GO

-- Creating table 'TestResultSet'
CREATE TABLE [dbo].[TestResultSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [TestCaseName] nvarchar(max)  NOT NULL,
    [Outcome] int  NOT NULL,
    [Duration] time  NOT NULL,
    [ErrorMessage] nvarchar(max)  NULL,
    [ErrorStackTrace] nvarchar(max)  NULL,
    [StartTime] datetimeoffset  NOT NULL,
    [EndTime] datetimeoffset  NOT NULL,
    [ComputerName] nvarchar(max)  NULL,
    [TestRunId] int  NOT NULL
);
GO

-- Creating table 'TestResultMessageSet'
CREATE TABLE [dbo].[TestResultMessageSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Category] nvarchar(max)  NOT NULL,
    [Text] nvarchar(max)  NOT NULL,
    [TestResultId] int  NOT NULL
);
GO

-- Creating table 'TestConfigurationSet'
CREATE TABLE [dbo].[TestConfigurationSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Description] nvarchar(max)  NULL,
    [VSBuildConfiguration] nvarchar(max)  NOT NULL,
    [Branch] nvarchar(max)  NOT NULL,
    [Type] int  NOT NULL,
    [TenantUrl] nvarchar(max)  NULL,
    [TestSiteUrl] nvarchar(max)  NOT NULL,
    [TestAuthentication_Id] int  NOT NULL
);
GO

-- Creating table 'TestAuthenticationSet'
CREATE TABLE [dbo].[TestAuthenticationSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Description] nvarchar(max)  NULL,
    [Type] int  NOT NULL,
    [AppOnly] bit  NOT NULL,
    [AppId] nvarchar(max)  NULL,
    [AppSecret] nvarchar(max)  NULL,
    [User] nvarchar(max)  NULL,
    [Domain] nvarchar(max)  NULL,
    [Password] nvarchar(max)  NULL,
    [CredentialManagerLabel] nvarchar(max)  NULL
);
GO

-- Creating table 'TestConfigurationPropertySet'
CREATE TABLE [dbo].[TestConfigurationPropertySet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Value] nvarchar(max)  NOT NULL,
    [TestConfigurationId] int  NOT NULL
);
GO

-- Creating table 'FileTrackingSet'
CREATE TABLE [dbo].[FileTrackingSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [TestDate] datetime  NOT NULL,
    [Build] nvarchar(max)  NOT NULL,
    [FileName] nvarchar(max)  NOT NULL,
    [FileHash] nvarchar(max)  NOT NULL,
    [FileChanged] bit  NOT NULL,
    [TestSiteUrl] nvarchar(max)  NOT NULL,
    [TestUser] nvarchar(max)  NULL,
    [TestAppId] nvarchar(max)  NULL,
    [TestComputerName] nvarchar(max)  NULL
);
GO

-- Creating table 'FileTrackingBaselineSet'
CREATE TABLE [dbo].[FileTrackingBaselineSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [FileName] nvarchar(max)  NOT NULL,
    [Build] nvarchar(max)  NOT NULL,
    [FileHash] nvarchar(max)  NOT NULL,
    [ChangeDate] datetime  NOT NULL,
    [FileContents] varbinary(max)  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'TestRunSet'
ALTER TABLE [dbo].[TestRunSet]
ADD CONSTRAINT [PK_TestRunSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'TestResultSet'
ALTER TABLE [dbo].[TestResultSet]
ADD CONSTRAINT [PK_TestResultSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'TestResultMessageSet'
ALTER TABLE [dbo].[TestResultMessageSet]
ADD CONSTRAINT [PK_TestResultMessageSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'TestConfigurationSet'
ALTER TABLE [dbo].[TestConfigurationSet]
ADD CONSTRAINT [PK_TestConfigurationSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'TestAuthenticationSet'
ALTER TABLE [dbo].[TestAuthenticationSet]
ADD CONSTRAINT [PK_TestAuthenticationSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'TestConfigurationPropertySet'
ALTER TABLE [dbo].[TestConfigurationPropertySet]
ADD CONSTRAINT [PK_TestConfigurationPropertySet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'FileTrackingSet'
ALTER TABLE [dbo].[FileTrackingSet]
ADD CONSTRAINT [PK_FileTrackingSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- Creating primary key on [Id] in table 'FileTrackingBaselineSet'
ALTER TABLE [dbo].[FileTrackingBaselineSet]
ADD CONSTRAINT [PK_FileTrackingBaselineSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [TestResultId] in table 'TestResultMessageSet'
ALTER TABLE [dbo].[TestResultMessageSet]
ADD CONSTRAINT [FK_TestResultTestResultMessage]
    FOREIGN KEY ([TestResultId])
    REFERENCES [dbo].[TestResultSet]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_TestResultTestResultMessage'
CREATE INDEX [IX_FK_TestResultTestResultMessage]
ON [dbo].[TestResultMessageSet]
    ([TestResultId]);
GO

-- Creating foreign key on [TestRunId] in table 'TestResultSet'
ALTER TABLE [dbo].[TestResultSet]
ADD CONSTRAINT [FK_TestRunTestResult]
    FOREIGN KEY ([TestRunId])
    REFERENCES [dbo].[TestRunSet]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_TestRunTestResult'
CREATE INDEX [IX_FK_TestRunTestResult]
ON [dbo].[TestResultSet]
    ([TestRunId]);
GO

-- Creating foreign key on [TestConfigurationId] in table 'TestRunSet'
ALTER TABLE [dbo].[TestRunSet]
ADD CONSTRAINT [FK_TestConfigurationTestRun]
    FOREIGN KEY ([TestConfigurationId])
    REFERENCES [dbo].[TestConfigurationSet]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_TestConfigurationTestRun'
CREATE INDEX [IX_FK_TestConfigurationTestRun]
ON [dbo].[TestRunSet]
    ([TestConfigurationId]);
GO

-- Creating foreign key on [TestAuthentication_Id] in table 'TestConfigurationSet'
ALTER TABLE [dbo].[TestConfigurationSet]
ADD CONSTRAINT [FK_TestConfigurationTestAuthentication]
    FOREIGN KEY ([TestAuthentication_Id])
    REFERENCES [dbo].[TestAuthenticationSet]
        ([Id])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_TestConfigurationTestAuthentication'
CREATE INDEX [IX_FK_TestConfigurationTestAuthentication]
ON [dbo].[TestConfigurationSet]
    ([TestAuthentication_Id]);
GO

-- Creating foreign key on [TestConfigurationId] in table 'TestConfigurationPropertySet'
ALTER TABLE [dbo].[TestConfigurationPropertySet]
ADD CONSTRAINT [FK_TestConfigurationTestConfigurationProperty]
    FOREIGN KEY ([TestConfigurationId])
    REFERENCES [dbo].[TestConfigurationSet]
        ([Id])
    ON DELETE CASCADE ON UPDATE NO ACTION;
GO

-- Creating non-clustered index for FOREIGN KEY 'FK_TestConfigurationTestConfigurationProperty'
CREATE INDEX [IX_FK_TestConfigurationTestConfigurationProperty]
ON [dbo].[TestConfigurationPropertySet]
    ([TestConfigurationId]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------