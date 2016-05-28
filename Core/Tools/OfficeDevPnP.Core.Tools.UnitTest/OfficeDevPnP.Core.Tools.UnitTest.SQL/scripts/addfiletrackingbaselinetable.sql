
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 12/29/2015 20:23:57
-- Generated from EDMX file: C:\GitHub\BertPnPSitesCore\Core\Tools\OfficeDevPnP.Core.Tools.UnitTest\OfficeDevPnP.Core.Tools.UnitTest.PnPBuildExtensions\SQL\TestModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [PnPTestAutomation];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
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


-- Creating primary key on [Id] in table 'FileTrackingBaselineSet'
ALTER TABLE [dbo].[FileTrackingBaselineSet]
ADD CONSTRAINT [PK_FileTrackingBaselineSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO
