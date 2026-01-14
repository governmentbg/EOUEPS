CREATE TABLE [dbo].[TransactionRequestIdentifiers](
	[TransactionRequestIdentifierId] [int] IDENTITY(1,1) NOT NULL,
	[Counter] [int] NOT NULL,
 CONSTRAINT [PK_TransactionRequestIdentifiers] PRIMARY KEY CLUSTERED 
(
	[TransactionRequestIdentifierId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[TransactionRequestIdentifiers] ADD  DEFAULT ((0)) FOR [Counter]