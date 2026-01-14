CREATE TABLE [dbo].[ObligationStatuses](
	[ObligationStatusId] [int] IDENTITY(1,1) NOT NULL,
	[Alias] [nvarchar](50) NOT NULL,
	[StatusText] [nvarchar](50) NOT NULL,
 CONSTRAINT [PK_ObligationStatuses] PRIMARY KEY CLUSTERED 
(
	[ObligationStatusId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]