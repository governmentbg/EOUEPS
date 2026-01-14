ALTER TABLE [dbo].[PaymentRequests]
	ADD [RedirectUrl] NVARCHAR(1000) NULL,

UPDATE [dbo].[GlobalValues]
	SET Value= '31'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2022-05-16_Version31', NULL, GETDATE());
