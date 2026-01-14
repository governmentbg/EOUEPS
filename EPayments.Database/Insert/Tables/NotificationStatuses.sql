PRINT 'Insert NotificationStatuses'
GO

SET IDENTITY_INSERT [NotificationStatuses] ON

INSERT INTO [NotificationStatuses]
    ([NotificationStatusId], [Alias], [Name], [Description])
VALUES
    (1                       , N'Pending'       , N'Чака за изпращане', NULL),
    (2                       , N'Sent'          , N'Изпратен', NULL),
    (3                       , N'Failure'       , N'Неуспешно изпращане', NULL),
    (4                       , N'Terminated'    , N'Прекратено поради невъзможност за изпращане', NULL)
GO

SET IDENTITY_INSERT [NotificationStatuses] OFF
