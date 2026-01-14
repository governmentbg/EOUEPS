CREATE TABLE [dbo].[Certificates](
	[CertificateId] [int] IDENTITY(1,1) NOT NULL,
	[CertificateFile] [varbinary](max) NULL,
	[CertificateSubject] [nvarchar](max) NULL,
	[CertificateIssuer] [nvarchar](max) NULL,
	[CertificateThumbprint] [nvarchar](200) NULL,
 CONSTRAINT [PK_Certificates] PRIMARY KEY CLUSTERED 
(
	[CertificateId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY],
 CONSTRAINT [UK_Certificates_Thumbprint] UNIQUE NONCLUSTERED 
(
	[CertificateThumbprint] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]