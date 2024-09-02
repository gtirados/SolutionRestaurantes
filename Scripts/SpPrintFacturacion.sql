IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SpPrintFacturacion'
)
BEGIN
    DROP PROC [dbo].[SpPrintFacturacion];
END;
GO
/*
SpPrintFacturacion '01','3',9,'B'
exec SpPrintFacturacion '01','3  ',5,'B'
select * from facart where far_numfac = 3
select * from allog w
*/
CREATE PROC [dbo].[SpPrintFacturacion]
    @Codcia CHAR(2),
    @Serie CHAR(3),
    @nro INT,
    @fbg CHAR(1)
AS
SET NOCOUNT ON;

DECLARE @PAGACON MONEY,
        @VUELTO MONEY;
SELECT @PAGACON = al.all_pagacon,
       @VUELTO = al.all_vuelto
FROM dbo.ALLOG al
WHERE al.ALL_CODCIA = @Codcia
      AND al.ALL_NUMSER = @Serie
      AND al.ALL_NUMFAC = @nro
      AND al.ALL_CODTRA = 2401
      AND al.ALL_FBG = @fbg
      AND al.ALL_SECUENCIA = 1;

SELECT CASE
           WHEN f.FAR_CANTIDAD_D IS NULL THEN
               f.FAR_CANTIDAD
           ELSE
               f.FAR_CANTIDAD_D
       END AS 'cant',
       CASE
           WHEN f.FAR_CANTIDAD_D IS NOT NULL THEN
               '1/2  '
           ELSE
               ''
       END + CASE
                 WHEN f.CAMBIOPRODUCTO IS NULL
                      OR f.CAMBIOPRODUCTO = '' THEN
                     a.ART_NOMBRE
                 ELSE
                     f.CAMBIOPRODUCTO
             END AS 'prod',
       f.FAR_PRECIO AS 'pre',
       (f.FAR_PRECIO * CASE
                           WHEN f.FAR_CANTIDAD_D IS NULL THEN
                               f.FAR_CANTIDAD
                           ELSE
                               f.FAR_CANTIDAD_D
                       END
       ) AS 'imp',
       dbo.FnDevuelveMozo(@Codcia, FAR_CODVEN) AS 'mozo',
       'Mesa: ' + FAR_OC AS 'Mesa',
       dbo.FnDevuelveCaracteristica(
                                       f.FAR_CODCIA,
                                       f.FAR_FECHA,
                                       f.FAR_NUMFAC_C,
                                       f.FAR_NUMSER_C,
                                       f.FAR_NUMSEC - 1,
                                       f.FAR_CODART
                                   ) AS 'carac',
       ISNULL(f.FAR_TOT_FLETE, 0) AS 'FLETE',
       ISNULL(f.FAR_TOT_DESCTO, 0) AS 'descuento',
       ISNULL(@PAGACON, 0) AS 'PAGACON',
       ISNULL(@VUELTO, 0) AS 'VUELTO',
       f.FAR_ICBPER AS 'ICBPER',
       (
           SELECT TOP 1
                  qr.CODIGOQR
           FROM dbo.DOCUMENTOS_QR qr
           WHERE qr.CODCIA = @Codcia
                 AND qr.NUMERO = f.FAR_NUMFAC
                 AND qr.FBG = f.FAR_FBG
                 AND qr.NSERIE = f.FAR_NUMSER
       ) AS 'codigoqr',
       f.FAR_NUMFAC_C AS 'COMANDA',
             (
           SELECT TOP 1
                  SUM(a2.ALL_SERVICIO)
           FROM dbo.ALLOG a2
           WHERE a2.ALL_CODCIA = f.FAR_CODCIA
                 AND RTRIM(LTRIM(a2.ALL_NUMSER)) = RTRIM(LTRIM(f.FAR_NUMSER))
                 AND a2.ALL_NUMFAC = f.FAR_NUMFAC
                 and a2.ALL_FBG = f.FAR_FBG
       ) AS 'servicio',
       (
           SELECT TOP 1
                  a2.ALL_SUBTRA
           FROM dbo.ALLOG a2
           WHERE a2.ALL_CODCIA = f.FAR_CODCIA
                 AND RTRIM(LTRIM(a2.ALL_NUMSER)) = RTRIM(LTRIM(f.FAR_NUMSER))
                 AND a2.ALL_NUMFAC = f.FAR_NUMFAC and a2.ALL_FBG = f.FAR_FBG
       ) AS 'FPAGO'
FROM FACART f
    INNER JOIN ARTI a
        ON f.FAR_CODART = a.ART_KEY
           AND f.FAR_CODCIA = a.ART_CODCIA
--inner join ALLOG al on f.FAR_CODCIA = al.ALL_CODCIA 
--                    and f.FAR_NUMSER = al.ALL_NUMSER 
--                    and f.FAR_NUMFAC = al.ALL_NUMFAC and al.ALL_CODTRA=2401
--                    AND f.FAR_FBG = al.ALL_FBG
WHERE f.FAR_CODCIA = @Codcia
      AND f.FAR_NUMSER = @Serie
      AND f.FAR_NUMFAC = @nro
      AND f.FAR_FBG = @fbg;





GO