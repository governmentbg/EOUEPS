BEGIN TRANSACTION
BEGIN TRY

-- create new tables

CREATE TABLE [dbo].[ObligationStatuses] (
	[ObligationStatusId]			INT 				NOT NULL IDENTITY,
	[Alias]                         NVARCHAR(50)        NOT NULL,
	[StatusText]					NVARCHAR(50)		NOT NULL,
CONSTRAINT [PK_ObligationStatuses]						PRIMARY KEY([ObligationStatusId])
);

CREATE TABLE [dbo].[PaymentRequestObligationLogs] (
	[PaymentRequestObligationLogsId]								INT				IDENTITY,
	[PaymentRequestId]												INT				NOT NULL,
	[ObligationStatusId]											INT				NOT NULL,
	[ChangeDate]													DATETIME2		NOT NULL,
CONSTRAINT [PK_PaymentRequestObligationLogsId]						PRIMARY KEY CLUSTERED ([PaymentRequestObligationLogsId]), 
CONSTRAINT [FK_PaymentRequestObligationLogs_PaymentRequests]		FOREIGN KEY([PaymentRequestId]) REFERENCES [dbo].[PaymentRequests] ([PaymentRequestId]),
CONSTRAINT [FK_PaymentRequestObligationLogs_ObligationStatuses]		FOREIGN KEY([ObligationStatusId]) REFERENCES [dbo].[ObligationStatuses] ([ObligationStatusId])
);

-- insert new data
SET IDENTITY_INSERT [dbo].[ObligationStatuses] ON

INSERT INTO [dbo].[ObligationStatuses] ([ObligationStatusId], [Alias], [StatusText])
VALUES
		(1, 'Asked', 'Заявено'),
		(2, 'Ordered', 'Наредено'),
		(3, 'IrrevocableOrder', 'Неотменимо нареждане'),
		(4, 'Canceled', 'Отказано плащане'),
		(5, 'Paid', 'Платено по сметка на ДАЕУ'),
		(6, 'ForDistribution', 'За разпределение'),
		(7, 'CheckedAccount', 'Заверена сметка на администрация');

SET IDENTITY_INSERT [dbo].[ObligationStatuses] OFF

-- update existing tables
ALTER TABLE [dbo].[PaymentRequests]
ADD [ObligationStatusId] INT NULL
CONSTRAINT [FK_PaymentRequests_ObligationStatuses] FOREIGN KEY ([ObligationStatusId]) REFERENCES [dbo].[ObligationStatuses] ([ObligationStatusId]);

-- update database values
UPDATE [dbo].[GlobalValues]
	SET Value= '13'
WHERE 
	[key] = 'DatabaseVersion';

INSERT INTO [dbo].[DatabaseUpdates] 
	([ScriptName], [Notes], [UpdateDate]) 
VALUES 
	( '2020-10-28_Version13', NULL, GETDATE());

COMMIT TRANSACTION
END TRY
BEGIN CATCH
ROLLBACK TRANSACTION
DECLARE @ErrorMessage NVARCHAR(MAX) = 'Error occured! Transaction for version 2020-10-28_Version13 has been rollback. Error message: ' + ERROR_MESSAGE();
RAISERROR(@ErrorMessage,16,1)
END CATCH
GO