IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'USP_COMPROBANTE_ELIMINAPAGOS'
)
BEGIN
    DROP PROC [dbo].[USP_COMPROBANTE_ELIMINAPAGOS];
END;
GO
/*

*/
CREATE PROC [dbo].[USP_COMPROBANTE_ELIMINAPAGOS]
    @CODCIA CHAR(2),
    @TIPODOCTO CHAR(1),
    @SERIE VARCHAR(3),
    @NUMERO BIGINT,
    @CURRENTUSER VARCHAR(20)
WITH ENCRYPTION
AS
BEGIN
    SET NOCOUNT ON;

    UPDATE dbo.COMPROBANTE_FORMAPAGO
    SET FE_DELETE = GETDATE(),
        CU_DELETE = @CURRENTUSER,
        ACTIVO = 0
    WHERE CODCIA = @CODCIA
          AND TIPODOCTO = @TIPODOCTO
          AND SERIE = @SERIE
          AND NUMERO = @NUMERO;



END;
GO