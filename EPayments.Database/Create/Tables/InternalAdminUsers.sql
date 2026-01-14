PRINT 'InternalAdminUsers'
GO

CREATE TABLE [dbo].[InternalAdminUsers](
    [InternalAdminUserId]                   INT                 NOT NULL IDENTITY,    
    [Name]                                  NVARCHAR(100)       NOT NULL,
	[Egn]									NVARCHAR(10)        NOT NULL UNIQUE,
	[IsSuperadmin]                          BIT                 NOT NULL,
    [IsActive]                              BIT                 NOT NULL,
	[CreateDate]                            DATETIME2           NOT NULL,

CONSTRAINT [PK_InternalAdminUsers]                              PRIMARY KEY ([InternalAdminUserId])
);
