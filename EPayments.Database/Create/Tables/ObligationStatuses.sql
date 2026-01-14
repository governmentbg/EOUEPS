PRINT 'ObligationStatuses'
GO

CREATE TABLE [dbo].[ObligationStatuses] (
	[ObligationStatusId]			INT 				NOT NULL IDENTITY,
	[Alias]                         NVARCHAR(50)        NOT NULL,
	[StatusText]					NVARCHAR(50)		NOT NULL,
CONSTRAINT [PK_ObligationStatuses]						PRIMARY KEY([ObligationStatusId])
);