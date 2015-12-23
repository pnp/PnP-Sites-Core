
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 12/22/2015 11:12:56
-- Generated from EDMX file: C:\GitHub\BertPnPSitesCore\Core\Tools\OfficeDevPnP.Core.Tools.UnitTest\OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions\SQL\TestModel.edmx
-- --------------------------------------------------
-- Manual cleaned to only create the FileTrackingSet

SET QUOTED_IDENTIFIER OFF;
GO
USE [PnPTestAutomation];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
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

-- Creating primary key on [Id] in table 'FileTrackingSet'
ALTER TABLE [dbo].[FileTrackingSet]
ADD CONSTRAINT [PK_FileTrackingSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------