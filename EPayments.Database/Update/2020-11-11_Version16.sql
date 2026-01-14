BEGIN TRANSACTION

BEGIN TRY

-- Users

ALTER TABLE [dbo].Users
	ADD [StatusObligationNotifications] bit NOT NULL DEFAULT(0) WITH VALUES;

-- Updated Database version

UPDATE [dbo].[GlobalValues]
	SET Value= '16'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-11-11_Version16', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-11-11_Version16 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO