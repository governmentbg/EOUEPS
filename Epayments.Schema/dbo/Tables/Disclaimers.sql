CREATE TABLE [dbo].[Disclaimers](
	[DisclaimerId] [int] IDENTITY(1,1) NOT NULL,
	[UserId] [int] NOT NULL,
	[CertificateId] [int] NOT NULL,
	[DisclaimerTemplateId] [int] NOT NULL,
	[SignedXml] [nvarchar](max) NOT NULL,
	[SignDate] [datetime] NOT NULL,
 CONSTRAINT [PK_Disclaimers] PRIMARY KEY CLUSTERED 
(
	[DisclaimerId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO
ALTER TABLE [dbo].[Disclaimers]  WITH CHECK ADD  CONSTRAINT [FK_Disclaimers_Certificates] FOREIGN KEY([CertificateId])
REFERENCES [dbo].[Certificates] ([CertificateId])
GO

ALTER TABLE [dbo].[Disclaimers] CHECK CONSTRAINT [FK_Disclaimers_Certificates]
GO
ALTER TABLE [dbo].[Disclaimers]  WITH CHECK ADD  CONSTRAINT [FK_Disclaimers_DisclaimerTemplates] FOREIGN KEY([DisclaimerTemplateId])
REFERENCES [dbo].[DisclaimerTemplates] ([DisclaimerTemplateId])
GO

ALTER TABLE [dbo].[Disclaimers] CHECK CONSTRAINT [FK_Disclaimers_DisclaimerTemplates]
GO
ALTER TABLE [dbo].[Disclaimers]  WITH CHECK ADD  CONSTRAINT [FK_Disclaimers_Users] FOREIGN KEY([UserId])
REFERENCES [dbo].[Users] ([UserId])
GO

ALTER TABLE [dbo].[Disclaimers] CHECK CONSTRAINT [FK_Disclaimers_Users]