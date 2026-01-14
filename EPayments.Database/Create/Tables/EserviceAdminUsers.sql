PRINT 'EserviceAdminUsers'
GO

CREATE TABLE [dbo].[EserviceAdminUsers](
    [EserviceAdminUserId]                   INT                 NOT NULL IDENTITY,
    [Username]                              NVARCHAR(100)       NOT NULL UNIQUE,
    [PasswordHash]                          NVARCHAR(MAX)       NOT NULL,
    [PasswordSalt]                          NVARCHAR(MAX)       NOT NULL,
    [IpAddressesForAccess]                  NVARCHAR(MAX)       NULL,
    [Name]                                  NVARCHAR(100)       NOT NULL,
    [Email]                                 NVARCHAR(100)       NULL,
    [InsufficientAmountNotifications]       BIT                 NOT NULL,
    [OverpaidAmountNotifications]           BIT                 NOT NULL,
    [ReferencedNotifications]               BIT                 NOT NULL,
    [NotReferencedNotifications]            BIT                 NOT NULL,
    [IsActive]                              BIT                 NOT NULL,
	[ReferringEserviceClientId]             INT                 NULL,

CONSTRAINT [PK_EserviceAdminUsers]                              PRIMARY KEY ([EserviceAdminUserId]),
CONSTRAINT [FK_EserviceAdminUsers_EserviceClients]              FOREIGN KEY ([ReferringEserviceClientId])     REFERENCES [dbo].[EserviceClients] ([EserviceClientId])
);



