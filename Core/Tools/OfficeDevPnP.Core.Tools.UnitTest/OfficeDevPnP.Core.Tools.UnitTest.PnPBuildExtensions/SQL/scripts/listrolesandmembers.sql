WITH RoleMembers (member_principal_id, role_principal_id) 
AS 
(
  SELECT 
   rm1.member_principal_id, 
   rm1.role_principal_id
  FROM sys.database_role_members rm1 
   UNION ALL
  SELECT 
   d.member_principal_id, 
   rm.role_principal_id
  FROM sys.database_role_members rm 
   INNER JOIN RoleMembers AS d 
   ON rm.member_principal_id = d.role_principal_id
)
select distinct rp.name as database_role, mp.name as database_userl
from RoleMembers drm
  join sys.database_principals rp on (drm.role_principal_id = rp.principal_id)
  join sys.database_principals mp on (drm.member_principal_id = mp.principal_id)
order by rp.name