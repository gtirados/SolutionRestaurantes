IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPCARGARCOMANDADELIVERY'
)
BEGIN
    DROP PROC [dbo].[SPCARGARCOMANDADELIVERY];
END;
GO
/*
exec SPCARGARCOMANDADELIVERY '01','100',16,'20240625'
*/
CREATE PROCEDURE [dbo].[SPCARGARCOMANDADELIVERY]
    @CodCia CHAR(2) ,
    @NUMSER CHAR(3) ,
    @NUMFAC BIGINT ,
    @Fecha DATETIME 

--With encryption
AS --Obtener la ultima comanda de acuerdo al valor maximo de ped_numfac

    SET NOCOUNT ON 
    
        DECLARE @ICBPER DECIMAL(8,2)
    
    SELECT TOP 1 @ICBPER = COALESCE(G.GEN_ICBPER,0) FROM dbo.GENERAL g
    
       DECLARE @TBLICBPER TABLE
                (
                  CODART BIGINT ,
                  ICBPER MONEY
                )
                
    
     INSERT  INTO @TBLICBPER
                    ( CODART ,
                      ICBPER
                    )
                    SELECT  p.PA_CODPA ,
                           SUM( p.PA_PROM*@ICBPER) AS 'Importe'
                    FROM    dbo.PAQUETES p
                            INNER JOIN dbo.ARTI art ON P.PA_CODCIA = art.ART_CODCIA
                                                       AND P.PA_CODART = art.art_key
                    WHERE   PA_CODPA IN (
                            SELECT  A2.ART_KEY
                            FROM    pedidos p
                                    INNER JOIN dbo.ARTI a2 ON p.PED_CODCIA = a2.ART_CODCIA
                                                              AND p.PED_CODART = a2.ART_KEY
                                    INNER JOIN ALLOG A ON p.PED_TRANSP = a.ALL_NUMOPER
                                                          AND A.ALL_CODCIA = @CODCIA
                                                          AND A.ALL_FECHA_DIA = @FECHA
                                                          AND A.ALL_FLAG_EXT = 'N'
                            WHERE  
                            -- ped_codclie = @CodMesa
--                                    AND 
                                    ped_codcia = @CodCia
                                    AND ped_fecha = @Fecha
                                    AND ped_estado = 'N'
                                    AND ped_situacion <> 'A'
                                    AND ped_fac <> ped_Cantidad
                                    AND ped_numfac = @NUMFAC
                                    AND a2.ART_FLAG_STOCK = 'C' )
                            AND art.ART_CALIDAD = 0 GROUP BY p.PA_CODPA
                            


    SELECT  a.ALL_NUMFAC AS numfac ,
            ped_numfac ,
            ped_numser ,
            ped_codven ,
            dbo.FnDevuelveMozo(@CodCia, PED_CODVEN) AS 'mozo' ,
            Ped_CodArt AS 'CodPlato' ,
            dbo.FnDevuelvePlato(@codcia, Ped_CodArt) AS 'Plato' ,
            ped_oferta AS 'Detalle' ,
            Ped_Precio AS 'Precio' ,
            Ped_Cantidad AS 'Cantidad' ,
            Ped_SubTotal AS 'Importe' ,
            ped_numsec AS 'sec' ,
            ped_aprobado AS 'apro' ,
            ped_canaten AS 'aten' ,
            ped_cta AS 'cuenta' ,
            P.PED_CLIENTE AS 'CLIENTE' ,
            P.PED_COMENSALES AS 'COMENSALES'
			,CASE (SELECT TOP 1 COALESCE(a2.ART_BOLSAS,0) FROM dbo.ARTI a2 WHERE a2.ART_CODCIA = p.PED_CODCIA AND a2.ART_KEY = p.PED_CODART) WHEN 0 THEN  0 ELSE 1 END AS 'ICBPER'
             --,case WHEN (SELECT COALESCE(XA.ART_CALIDAD,1) FROM dbo.ARTI xa WHERE XA.ART_CODCIA = @CODCIA AND XA.ART_KEY = p.PED_CODART) = 0 then 1 else 0 END AS 'ICBPER'
                    ,@ICBPER AS 'GEN_ICBPER',
                     COALESCE(t.ICBPER*ped_cantidad, 0) AS 'COMBO_ICBPER'
					 ,(SELECT TOP 1 COALESCE(a3.ART_BOLSAS,0) FROM dbo.ARTI a3 WHERE a3.ART_CODCIA = P.PED_CODCIA AND a3.ART_KEY = P.PED_CODART) AS 'BOLSAS'
    FROM    pedidos p
            INNER JOIN ALLOG A ON p.PED_TRANSP = a.ALL_NUMOPER
                                  AND A.ALL_CODCIA = @CODCIA
                                  AND A.ALL_FECHA_DIA = @FECHA
                                  AND A.ALL_FLAG_EXT = 'N'
                                  LEFT JOIN @TBLICBPER t ON P.PED_CODART = t.CODART
	--inner join allog on pedidos.ped_codcia = allog.all_codcia and pedidos.ped_fecha = allog.all_fecha_dia
	--and pedidos.ped_codart = allog.all_codclie
    WHERE   ped_codcia = @CodCia
            AND ped_fecha = @Fecha
            AND ped_estado = 'N'
            AND ped_situacion <> 'A' --and ped_fac <> ped_Cantidad
            AND ped_numfac = @NUMFAC
            AND P.PED_NUMSER = @NUMSER

      
 SELECT  ISNULL(PC.DIRECCION,'') AS 'DIRECCION' ,
            ISNULL(C.CLI_NOMBRE,'') AS 'CLIENTE' ,
            ISNULL(PC.IDCLIENTE,-1) AS 'IDECLIENTE',
            ISNULL(PC.PAGO,0) AS 'PAGO',
            ISNULL(PC.OBS,'') AS 'OBS',
            ISNULL(c.CLI_RUC_ESPOSO,'') AS 'ruc',
            ISNULL(PC.IDZONA,0) AS 'IDZ',
            COALESCE(c.CLI_RUC_ESPOSA,'') AS 'dni',
             COALESCE(PC.RECOJO_TIENDA,0) AS 'RECOJO',
             COALESCE(PC.IDDIRECCION,0) AS 'IDDIR'
    FROM    dbo.PEDIDOS_CABECERA pc
            LEFT JOIN dbo.CLIENTES c ON PC.IDCLIENTE = C.CLI_CODCLIE
                                        AND PC.CODCIA = C.CLI_CODCIA
                                        AND C.CLI_ESTADO = 'A' AND c.CLI_CP='C'
    WHERE   PC.CODCIA = @CodCia
            AND PC.NUMFAC = @NUMFAC
            AND PC.NUMSER = @NUMSER
            
            
            

GO