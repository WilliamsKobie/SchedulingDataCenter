USE [FamilyLiteracy.mdf]
GO
/****** Object:  StoredProcedure [dbo].[SearchGuardianFirstName]    Script Date: 3/28/2015 10:08:36 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
ALTER PROCEDURE [dbo].[SearchGuardianFirstName]
	(@firstname varchar(50))
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	Select [First Name]  FROM [dbo].[GuardianProfile] Where [First Name] Like @firstname + '%'
END
