IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_EMPRESA_IGV'
)
BEGIN
    DROP PROC [dbo].[USP_EMPRESA_IGV];
END;
GO
/*
USP_EMPRESA_IGV '01'
USP_EMPRESA_IGV '02'
*/
CREATE PROC [dbo].[USP_EMPRESA_IGV] @CODCIA CHAR(2)
AS
SET NOCOUNT ON;
SELECT p.PAR_IGV,
       COALESCE(p.PAR_ICBPER, 0) AS 'icbper'
FROM PARGEN p
WHERE PAR_CODCIA = @CODCIA;
GO