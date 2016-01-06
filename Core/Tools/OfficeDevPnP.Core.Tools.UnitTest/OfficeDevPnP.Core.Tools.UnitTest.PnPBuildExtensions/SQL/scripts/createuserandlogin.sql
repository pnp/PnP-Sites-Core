-- in master db
create login PnP with password = 'pwd'
create login PnPReader with password = 'pwd'

-- in PnP db
CREATE USER [PnP] FOR LOGIN [PnP] WITH DEFAULT_SCHEMA = dbo
CREATE USER [PnPReader]	FOR LOGIN [PnPReader] WITH DEFAULT_SCHEMA = dbo

-- Add user to the database owner role
EXEC sp_addrolemember 'db_datareader', 'pnp'
EXEC sp_addrolemember 'db_datawriter', 'pnp'
EXEC sp_addrolemember 'PnPLimitedReader', 'PnPReader'




