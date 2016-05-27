Alter table [dbo].[UsersSet] add IsCoreMember BIT NOT NULL DEFAULT 0
Alter table [dbo].[UsersSet] add SendTestResults BIT NOT NULL DEFAULT 0
Alter table [dbo].[UsersSet] add Email nvarchar(max)  NULL
Alter table [dbo].[UsersSet] add IsEmailVerified BIT NOT NULL DEFAULT 0
