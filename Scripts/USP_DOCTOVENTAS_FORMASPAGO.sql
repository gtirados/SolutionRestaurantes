--exec [dbo].[USP_DOCTOVENTAS_FORMASPAGO] '01','B','3','4416'

/*  
[dbo].[USP_DOCTOVENTAS_FORMASPAGO] '01','B','2',2  
*/  
alter PROCEDURE [dbo].[USP_DOCTOVENTAS_FORMASPAGO]  
    @CODCIA CHAR(2) ,  
    @FBG CHAR(1) ,  
    @SERIE VARCHAR(3) ,  
    @NUMERO INT  
AS -- statements  
  SET NOCOUNT ON
  

    --SELECT  ALL_NUMSER AS 'SERIE' ,  
    --        ALL_NUMFAC AS 'NUMERO' ,  
    --        ALL_FECHA_DIA AS 'fecha' ,  
    --        RTRIM(LTRIM(ST.SUT_DESCRIPCION)) AS 'FORMAPAGO' ,  
    --        A.ALL_IMPORTE_AMORT AS 'IMPORTE' ,  
    --        A.ALL_NETO AS 'TOTAL',A.ALL_NUMOPER AS 'NUMOPER',A.ALL_FBG AS 'FBG'  
    --FROM    dbo.ALLOG a  
    --        INNER JOIN dbo.SUB_TRANSA st ON A.ALL_SECUENCIA = ST.SUT_SECUENCIA  
    --                                        AND ST.SUT_CODTRA = 2401  
    --WHERE   a.ALL_CODCIA = @CODCIA  
    --        AND ALL_CODTRA = 2401  
    --        AND ALL_FBG = @FBG  
    --        AND ALL_NUMSER = @SERIE  
    --        AND ALL_NUMFAC = @NUMERO  
  
  
  select RTRIM(LTRIM(SERIE)) AS 'SERIE',NUMERO,
  (select top 1 ALL_FECHA_DIA from ALLOG where ALL_NUMSER= @SERIE and ALL_NUMFAC = @NUMERO and ALL_FBG = TIPODOCTO) as FECHA
           ,RTRIM(LTRIM(st.SUT_DESCRIPCION)) as 'FORMAPAGO',MONTO as 'IMPORTE',
           (select SUM(ALL_NETO) from ALLOG where ALL_NUMSER= @SERIE and ALL_NUMFAC = @NUMERO and ALL_FBG = TIPODOCTO) AS 'TOTAL',
           TIPODOCTO AS 'FBG',
           IDFORMAPAGO
           ,correlativo
  from COMPROBANTE_FORMAPAGO 
   INNER JOIN dbo.SUB_TRANSA st ON idformapago = st.sut_secuencia and st.sut_codtra=2401
   where TIPODOCTO = @FBG AND SERIE = @SERIE AND NUMERO = @NUMERO and COMPROBANTE_FORMAPAGO.CODCIA = @CODCIA;