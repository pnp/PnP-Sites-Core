CREATE TABLE [dbo].[UsersSet] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [UPN] nvarchar(max)  NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [IsAdmin] bit  NOT NULL
);