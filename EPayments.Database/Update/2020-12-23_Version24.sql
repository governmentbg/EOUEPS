BEGIN TRANSACTION

BEGIN TRY

ALTER TABLE [dbo].[DistributionRevenuePayments]
	ADD [BoricaTransactionId]												INT					NOT NULL,
		[DitribtutionError]													NVARCHAR(500)		NULL,
	CONSTRAINT [FK_DistributionRevenuePayments_BoricaTransactions]								FOREIGN KEY ([BoricaTransactionId]) REFERENCES [dbo].[BoricaTransactions]([BoricaTransactionId]);

UPDATE [dbo].[GlobalValues]
	SET Value= '24'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-12-23_Version24', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-12-23_Version24 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO