BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '5';
DECLARE @ScriptName NVARCHAR(MAX) = '2016-01-18_Version5.sql';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

--Set NACID VPOS production values

update VposClients
set PaymentRequestUrl=N'https://gate.borica.bg/boreps/registerTransaction'
where Alias=N'UniCredit'

update EserviceClients
set
	BoricaVposTerminalId=N'62160737',
	BoricaVposMerchantId=N'6210015685'
where Alias=N'Nacid'


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






