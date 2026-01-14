BEGIN TRANSACTION

BEGIN TRY

ALTER TABLE [dbo].[EserviceClients]
	ADD			[BudgetCode]		NVARCHAR(10)		NULL;

INSERT INTO [dbo].[GlobalValues]
(
	[Key],
	[Value],
	[ModifyDate]
)
VALUES
(
	'CVPosTransactionJobInvocationTime',
	null,
	GETDATE()
);


UPDATE [dbo].[GlobalValues]
	SET [Value] = '28'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2021-05-15_Version28', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2021-05-15_Version28 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO