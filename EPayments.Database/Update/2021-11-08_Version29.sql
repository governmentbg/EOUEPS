ALTER TABLE [dbo].[PaymentRequests]
	ADD [InitiatorId]									INT			NOT NULL DEFAULT(1) WITH VALUES,
	CONSTRAINT [FK_PaymentRequests_EServiceClients_InitiatorId]						FOREIGN KEY ([InitiatorId]) REFERENCES [dbo].[EserviceClients](EserviceClientId);

UPDATE [dbo].[PaymentRequests]
SET [InitiatorId] = [EserviceClientId]
-- Updated Database version

UPDATE [dbo].[GlobalValues]
	SET Value= '29'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2021-11-0_Version29', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-11-21_Version17 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO