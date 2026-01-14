CREATE VIEW [dbo].[vwDepartmentAisClients]
AS  
select 
ec.EserviceClientId,
d.DepartmentId,
d.Name as DepartmentName,
ec.AisName,
ec.ClientId,
ec.SecretKey,
ec.AccountBank as ServiceProviderBank,
ec.AccountBIC as serviceProviderBIC,
ec.AccountIBAN as serviceProviderIBAN
from EserviceClients ec
inner join Departments d on  d.DepartmentId=ec.DepartmentId