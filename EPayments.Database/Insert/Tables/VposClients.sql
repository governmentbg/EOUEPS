PRINT 'Insert [VposClients]'
GO

SET IDENTITY_INSERT [dbo].[VposClients] ON 
GO

INSERT [dbo].[VposClients] ([VposClientId], [Name], [Alias], [PaymentRequestUrl], [IsActive])
VALUES (1, N'ДСК (стар протокол)', N'Dsk', N'https://www.dskdirect.bg/e-commerce/default.aspx?xml_id=/bg-BG/.CardPayments/', 1)

INSERT [dbo].[VposClients] ([VposClientId], [Name], [Alias], [PaymentRequestUrl], [IsActive])
VALUES (2, N'УниКредит Виртуален ПОС', N'UniCredit', N'https://gatet.borica.bg/boreps/registerTransaction', 1)

INSERT [dbo].[VposClients] ([VposClientId], [Name], [Alias], [PaymentRequestUrl], [IsActive])
VALUES (3, N'FiBank Виртуален ПОС', N'FiBank', NULL, 1)

INSERT [dbo].[VposClients] ([VposClientId], [Name], [Alias], [PaymentRequestUrl], [IsActive])
VALUES (4, N'ДСК Виртуален ПОС (Ecomm)', N'DskEcomm', N'https://ecomtest.dskbank.bg', 1)


SET IDENTITY_INSERT [dbo].[VposClients] OFF
GO