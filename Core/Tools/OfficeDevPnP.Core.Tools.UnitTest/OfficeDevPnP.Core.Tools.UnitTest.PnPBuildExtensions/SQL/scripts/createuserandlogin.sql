-- in master db
create login PnPReader with password = 'pwd'

-- in PnP db
CREATE USER [PnPReader]	FOR LOGIN [PnPReader] WITH DEFAULT_SCHEMA = dbo

-- Add user to a role
EXEC sp_addrolemember 'PnPLimitedReader', 'PnPReader'




