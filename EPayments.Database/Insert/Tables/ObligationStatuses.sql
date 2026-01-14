PRINT 'Insert ObligationStatuses'
GO

SET IDENTITY_INSERT [dbo].[ObligationStatuses] ON;

INSERT INTO [dbo].[ObligationStatuses] ([ObligationStatusId], [Alias], [StatusText])
VALUES
		(1, 'Asked', 'Заявено'),
		(2, 'Ordered', 'Наредено'),
		(3, 'IrrevocableOrder', 'Неотменимо нареждане'),
		(4, 'Canceled', 'Отказано плащане'),
		(5, 'Paid', 'Платено по сметка на ДАЕУ'),
		(6, 'ForDistribution', 'За разпределение'),
		(7, 'CheckedAccount', 'Заверена сметка на администрация');

SET IDENTITY_INSERT [dbo].[ObligationStatuses] OFF;