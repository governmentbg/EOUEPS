CREATE TABLE [dbo].[Users](
	[UserId] [int] IDENTITY(1,1) NOT NULL,
	[Egn] [nvarchar](15) NOT NULL,
	[Email] [nvarchar](100) NULL,
	[RequestNotifications] [bit] NOT NULL,
	[StatusNotifications] [bit] NOT NULL,
	[LastCertificateId] [int] NULL,
	[LastDisclaimerTemplateId] [int] NULL,
	[AccessCodeNotifications] [bit] NOT NULL,
	[StatusObligationNotifications] [bit] NOT NULL,
 CONSTRAINT [PK_Users] PRIMARY KEY CLUSTERED 
(
	[UserId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UK_Users_Egn] UNIQUE NONCLUSTERED 
(
	[Egn] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UQ__Users__C1972A8EE5F6BFAF] UNIQUE NONCLUSTERED 
(
	[Egn] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[Users]  WITH CHECK ADD  CONSTRAINT [FK_Users_Certificates] FOREIGN KEY([LastCertificateId])
REFERENCES [dbo].[Certificates] ([CertificateId])
GO

ALTER TABLE [dbo].[Users] CHECK CONSTRAINT [FK_Users_Certificates]
GO
ALTER TABLE [dbo].[Users]  WITH CHECK ADD  CONSTRAINT [FK_Users_DisclaimerTemplates] FOREIGN KEY([LastDisclaimerTemplateId])
REFERENCES [dbo].[DisclaimerTemplates] ([DisclaimerTemplateId])
GO

ALTER TABLE [dbo].[Users] CHECK CONSTRAINT [FK_Users_DisclaimerTemplates]
GO
ALTER TABLE [dbo].[Users] ADD  CONSTRAINT [DF__Users__StatusObl__32767D0B]  DEFAULT ((0)) FOR [StatusObligationNotifications]