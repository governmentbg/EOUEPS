CREATE TABLE [dbo].[VposRedirects](
	[VposRedirectId] [int] IDENTITY(1,1) NOT NULL,
	[Gid] [uniqueidentifier] NOT NULL,
	[PaymentRequestIdentifier] [nvarchar](50) NULL,
	[OkUrl] [nvarchar](max) NULL,
	[CancelUrl] [nvarchar](max) NULL,
	[IP] [nvarchar](50) NULL,
	[LogDate] [datetime] NOT NULL,
 CONSTRAINT [PK_VposRedirects] PRIMARY KEY CLUSTERED 
(
	[VposRedirectId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
UNIQUE NONCLUSTERED 
(
	[Gid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]