CREATE TABLE [dbo].[DisclaimerTemplates](
	[DisclaimerTemplateId] [int] IDENTITY(1,1) NOT NULL,
	[DisclaimerText] [nvarchar](max) NOT NULL,
	[CreateDate] [datetime] NOT NULL,
	[IsActive] [bit] NOT NULL,
 CONSTRAINT [PK_DisclaimerTemplates] PRIMARY KEY CLUSTERED 
(
	[DisclaimerTemplateId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]