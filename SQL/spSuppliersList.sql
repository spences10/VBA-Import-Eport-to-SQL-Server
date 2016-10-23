USE [Northwind]
GO

/****** Object:  StoredProcedure [dbo].[spSuppliersList]    Script Date: 18/10/2016 19:51:37 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[spSuppliersList]

AS   

    SET NOCOUNT ON;  
    SELECT CompanyName, Id
    FROM Supplier
	ORDER BY CompanyName

GO

