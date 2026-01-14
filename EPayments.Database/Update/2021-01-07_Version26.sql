BEGIN TRANSACTION

BEGIN TRY

ALTER TABLE [dbo].[EserviceDeliveryNotifications]
	ADD [PersonUniqueIdentifier] NVARCHAR(20) NULL,
		[ResponseCodes] NVARCHAR(100) NULL;

ALTER TABLE [dbo].[EserviceDeliveryNotifications]
	DROP COLUMN [MessageText];

UPDATE [dbo].[GlobalValues]
	SET Value= '26'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2021-01-07_Version26', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2021-01-07_Version26 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO