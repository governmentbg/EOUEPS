CREATE TABLE [dbo].[DatabaseUpdates](
	[DatabaseUpdateId] [int] IDENTITY(1,1) NOT NULL,
	[ScriptName] [nvarchar](200) NOT NULL,
	[Notes] [nvarchar](max) NULL,
	[UpdateDate] [datetime2](7) NOT NULL,
 CONSTRAINT [PK_DatabaseUpdates] PRIMARY KEY CLUSTERED 
(
	[DatabaseUpdateId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]