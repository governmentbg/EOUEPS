PRINT 'Insert [GlobalValues]'
GO


INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'DatabaseVersion', N'Initial', GETDATE())

INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'EmailJobLastInvocationTime', NULL, GETDATE())

INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'EserviceNotificationJobLastInvocationTime', NULL, GETDATE())

INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'ExpiredRequestJobLastInvocationTime', NULL, GETDATE())

INSERT [dbo].[GlobalValues] ([Key], [Value], [ModifyDate])
VALUES (N'EventRegisterNotificationJobLastInvocationTime', NULL, GETDATE())
