
--Note: this adds the TestCategory_Id foreign key relationship with a key value than can be null, which allows one to create the relationship with existing data 
--Use addtestcategorytableandconstraints - part 2.sql once you've run this script and fixed the data

SET QUOTED_IDENTIFIER ON
SET ANSI_NULLS ON
SET ANSI_PADDING ON
SET ANSI_WARNINGS ON
SET ARITHABORT ON
SET NUMERIC_ROUNDABORT OFF
SET CONCAT_NULL_YIELDS_NULL ON
SET XACT_ABORT ON
GO


BEGIN TRANSACTION
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[TestCategorySet]
(
	[Id] [int] IDENTITY(1,1),
	[Name] [nvarchar](max) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL
)
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
ALTER TABLE [dbo].[TestConfigurationSet] ADD [TestCategory_Id] [int] 
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
ALTER TABLE [dbo].[TestCategorySet] ADD CONSTRAINT [PK_TestCategorySet] PRIMARY KEY CLUSTERED
(
	[Id] ASC
)
WITH (STATISTICS_NORECOMPUTE = OFF)
GO

CREATE NONCLUSTERED INDEX [IX_FK_TestCategoryTestConfiguration] ON [dbo].[TestConfigurationSet]
(
	[TestCategory_Id] ASC
)
WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF)
GO

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
ALTER TABLE [dbo].[TestConfigurationSet] WITH CHECK ADD CONSTRAINT [FK_TestCategoryTestConfiguration] FOREIGN KEY
(
	[TestCategory_Id]
)
REFERENCES [dbo].[TestCategorySet]
(
	[Id]
)
ON DELETE NO ACTION
ON UPDATE NO ACTION
GO


COMMIT TRANSACTION
GO


