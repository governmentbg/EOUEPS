PRINT 'Departments'
GO

CREATE TABLE [dbo].[Departments](
    [DepartmentId]      INT             NOT NULL IDENTITY,
    [Name]              NVARCHAR(200)   NOT NULL,
 CONSTRAINT [PK_Departments] PRIMARY KEY ([DepartmentId])
);

GO
