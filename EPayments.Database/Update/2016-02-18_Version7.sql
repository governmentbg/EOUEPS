BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '7';
DECLARE @ScriptName NVARCHAR(MAX) = '2016-02-18_Version7';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------


Update EserviceClients
set 
DskVposMerchantId=N'ffb2070bd0bb11e58734005056b679e2',
DskVposMerchantPassword=N'vpst34tcrdpmnt18'
where Alias=N'ArchimedLovech'

Update EserviceClients
set 
DskVposMerchantId=N'55b801d3d0bb11e58734005056b679e2',
DskVposMerchantPassword=N'st34asumvur6dfs'
where Alias=N'ArchimedStaraZagora'

Update EserviceClients
set 
DskVposMerchantId=N'b79d0933d0bb11e58734005056b679e2',
DskVposMerchantPassword=N'hger457dfghhtr56g'
where Alias=N'ArchimedHaskovo'

------

SET IDENTITY_INSERT [dbo].[Departments] ON 

INSERT [dbo].[Departments] ([DepartmentId], [Alias], [Name])
VALUES (5, N'NSI', N'Национален статистически институт')

SET IDENTITY_INSERT [dbo].[Departments] OFF 

-------

SET IDENTITY_INSERT [dbo].[EserviceClients] On

INSERT [dbo].[EserviceClients] ([EserviceClientId], [VposClientId], [Alias], [ClientId], [AisName], [ServiceName], [ServiceDescription], [AccountBank], [AccountBIC], [AccountIBAN], [SecretKey], [IsActive], [DskVposMerchantId], [DskVposMerchantPassword], [Gid], [IsAuthPassAuthorized], [DisableUniqueRequestConstraint], [DepartmentId])
VALUES (7, 1, N'NsiInfostat', N'nsi_infostat_af42a889-0621-4a27-975d-51516ff8b7c3', N'ИС "Инфостат" НСИ', N'', NULL, 'Банка ДСК', 'STSABGSF', '', N'S23LPNVUUPTYASBK9UQ5Y1Z79UEXD4PZIEHDQXEYZHIYP7YABUZQS1X98D2MAQTP6TGYNUPT647HZ4BGYQKITTSJGCEBBEC5EYA67KBM11REXVCNE52GUM27ZYM1MC98', 1, N'2ee50607d0bc11e58734005056b679e2', N'vrtlps1txmcnsrc16', '08bf5c16-5b65-475b-a8fe-f6dc89f7c62e', 0, 0, 5)

SET IDENTITY_INSERT [dbo].[EserviceClients] OFF

---------

Update EserviceClients
set 
BoricaVposTerminalId=N'62160735',
BoricaVposMerchantId=N'6210015673'
where Alias=N'Rta'


--------------------------------------------------------------
--Update database version
Update [dbo].[GlobalValues] 
SET [Value] = @DbVersion,
    [ModifyDate] = GETDATE()
WHERE [Key]='DatabaseVersion'
--Add DatabaseUpdate record
INSERT INTO [dbo].[DatabaseUpdates] (ScriptName, Notes, UpdateDate)
VALUES (@ScriptName, @Notes, GETDATE())
-------------------------------------------------------------
--End commands block
COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH






