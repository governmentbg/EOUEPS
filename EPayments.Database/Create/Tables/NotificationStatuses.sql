PRINT 'NotificationStatuses'
GO

CREATE TABLE [dbo].[NotificationStatuses](
    [NotificationStatusId]    INT             NOT NULL IDENTITY,
    [Alias]            NVARCHAR(50)    NOT NULL,
    [Name]             NVARCHAR(100)   NOT NULL,
    [Description]      NVARCHAR(MAX)   NULL,
 CONSTRAINT [PK_NotificationStatuses] PRIMARY KEY ([NotificationStatusId])
);

GO
