SELECT p.name, o.name, d.*
FROM sys.database_principals AS p
JOIN sys.database_permissions AS d ON d.grantee_principal_id = p.principal_id
JOIN sys.objects AS o ON o.object_id = d.major_id