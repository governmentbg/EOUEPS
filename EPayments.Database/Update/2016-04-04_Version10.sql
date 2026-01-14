BEGIN TRANSACTION
BEGIN TRY
--Begin commands block (all "GO" statements must be removed)
--------------------------------------------------------------
DECLARE @DbVersion NVARCHAR(MAX) = '10';
DECLARE @ScriptName NVARCHAR(MAX) = '2016-04-04_Version10';
DECLARE @Notes NVARCHAR(MAX) = NULL;
--------------------------------------------------------------

ALTER TABLE EserviceClients DROP COLUMN DisableUniqueRequestConstraint;

ALTER TABLE PaymentRequests ADD AisPaymentId NVARCHAR(MAX) NULL;

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






