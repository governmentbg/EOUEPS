BEGIN TRANSACTION

BEGIN TRY

CREATE TABLE [dbo].[TransactionRequestIdentifiers]
(
	[TransactionRequestIdentifierId]					INT						NOT NULL IDENTITY,
	[Counter]											INT						NOT NULL DEFAULT(0),
	CONSTRAINT [PK_TransactionRequestIdentifiers]								PRIMARY KEY ([TransactionRequestIdentifierId])
);

INSERT INTO [dbo].[TransactionRequestIdentifiers]
	([Counter])
VALUES
	(0);

UPDATE [dbo].[GlobalValues]
	SET Value= '23'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-12-21_Version23', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-12-21_Version23 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO