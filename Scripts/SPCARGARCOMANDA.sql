IF EXISTS
(
    SELECT TOP 1
        s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPCARGARCOMANDA'
)
BEGIN
    DROP PROC [dbo].[SPCARGARCOMANDA];
END;
GO

/*
exec SpCargarComanda '01','7B','2020-08-17 00:00:00',1
exec SpCargarComanda '01','16','20240625',1
*/
CREATE PROCEDURE [dbo].[SPCARGARCOMANDA]
    @CodCia CHAR(2) ,
    @CodMesa VARCHAR(10) ,
    @Fecha DATE ,
    @Fac BIT = NULL
--@NumSer char(3) out,
--@NumFac int out,
--@CodMozo int out,
--@Mozo varchar(50) out
--With encryption
AS --Obtener la ultima comanda de acuerdo al valor maximo de ped_numfac
--Consultar con el ing. si sabe una mejor forma de obtener el pedido actual de la mesa

    DECLARE @tt INT
/*    SELECT  @tt = ISNULL(MAX(ped_numfac), 0)
    FROM    pedidos
    WHERE   ped_codclie = @codmesa
            AND ped_codcia = @codcia
            AND ped_fecha = @fecha */
    SELECT TOP 1
            @tt = pc.NUMFAC
    FROM    dbo.PEDIDOS_CABECERA pc
    WHERE   pc.CODCIA = @CodCia
            AND CONVERT(VARCHAR(8), pc.FECHA, 112) = @Fecha
            AND pc.CODMESA = @CodMesa
            AND pc.FACTURADO = 0
    ORDER BY pc.NUMFAC DESC
    
    DECLARE @ICBPER DECIMAL(8, 2)
    
    SELECT TOP 1
            @ICBPER = COALESCE(G.GEN_ICBPER, 0)
    FROM    dbo.GENERAL g
	
    
     DECLARE @TBLICBPER TABLE
                (
                  CODART BIGINT ,
                  ICBPER MONEY
                )

    IF @fac IS NULL
        BEGIN

            SELECT  a.ALL_NUMFAC AS numfac ,
                    ped_numfac ,
                    ped_numser ,
                    ped_codven ,
                    --dbo.FnDevuelveMozo(@CodCia, PED_CODVEN) AS 'mozo' ,
                    V.VEM_NOMBRE AS 'MOZO' ,
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
                    P.PED_COMENSALES AS 'COMENSALES' ,
                    CAST(p.CANTADO AS INT) AS 'CANTADO' ,
                    CASE WHEN P.PED_ENVIAR_EN IS NULL THEN 'AHORA'
                         ELSE CAST(DATEDIFF(MINUTE, P.PED_FECHAREG,
                                            P.PED_ENVIAR_EN) AS VARCHAR(20))
                              + ' min'
                    END AS 'ENVIAR' ,
                    P.PED_FAMILIA2 AS 'FAM'
            FROM    pedidos p
                    INNER JOIN dbo.PEDIDOS_CABECERA pc ON P.PED_FECHA = PC.FECHA
                                                          AND P.PED_CODCIA = PC.CODCIA
                                                          AND P.PED_NUMSER = PC.NUMSER
                                                          AND P.PED_NUMFAC = PC.NUMFAC
                    INNER JOIN dbo.VEMAEST v ON PC.CODCIA = V.VEM_CODCIA
                                                AND PC.CODMOZO = V.VEM_CODVEN
                    INNER JOIN ALLOG A ON p.PED_TRANSP = a.ALL_NUMOPER
                                          AND A.ALL_CODCIA = @CODCIA
                                          AND A.ALL_FECHA_DIA = @FECHA
                                          AND A.ALL_FLAG_EXT = 'N'
	--inner join allog on pedidos.ped_codcia = allog.all_codcia and pedidos.ped_fecha = allog.all_fecha_dia
	--and pedidos.ped_codart = allog.all_codclie
            WHERE   ped_codclie = @CodMesa
                    AND ped_codcia = @CodCia
                    AND ped_fecha = @Fecha
                    AND ped_estado = 'N'
                    AND ped_situacion <> 'A' --and ped_fac <> ped_Cantidad
                    AND ped_numfac = @tt

        END
    ELSE
        BEGIN
         --   INSERT  INTO @TBLICBPER
         --           ( CODART ,
         --             ICBPER
         --           )
         --           SELECT  p.PA_CODPA ,
         --                  SUM( p.PA_PROM*@ICBPER) AS 'Importe'
         --           FROM    dbo.PAQUETES p
         --                   INNER JOIN dbo.ARTI art ON P.PA_CODCIA = art.ART_CODCIA
         --                                              AND P.PA_CODART = art.art_key
                                                   
         --           WHERE   PA_CODPA IN (
         --                   SELECT  A2.ART_KEY
         --                   FROM    pedidos p
         --                           INNER JOIN dbo.ARTI a2 ON p.PED_CODCIA = a2.ART_CODCIA
         --                                                     AND p.PED_CODART = a2.ART_KEY
         --                           INNER JOIN ALLOG A ON p.PED_TRANSP = a.ALL_NUMOPER
         --                                                 AND A.ALL_CODCIA = @CODCIA
         --                                                 AND A.ALL_FECHA_DIA = @FECHA
         --                                                 AND A.ALL_FLAG_EXT = 'N'
         --                   WHERE   ped_codclie = @CodMesa
         --                           AND ped_codcia = @CodCia
         --                           AND ped_fecha = @Fecha
         --                           AND ped_estado = 'N'
         --                           AND ped_situacion <> 'A'
         --                           AND ped_fac <> ped_Cantidad
         --                           AND ped_numfac = @tt
         --                           AND a2.ART_FLAG_STOCK = 'C' 
									--)
         --                   AND art.ART_CALIDAD = 0 GROUP BY p.PA_CODPA
                            

            SELECT 
--dbo.FnDevuelveNumOper(@CodCia,@Fecha,ped_transp,ped_codart) as numfac, 
                    a.ALL_NUMFAC AS numfac ,
                    Ped_CodArt AS 'CodPlato' ,
                    dbo.FnDevuelvePlato(@codcia, Ped_CodArt) AS 'Plato' ,
                    Ped_Precio AS 'Precio' ,
                    --( ped_cantidad - ped_FAC ) AS 'CantTotal' ,
                    CASE WHEN CANTIDAD_DELIVERY IS NULL
                         THEN PED_CANTIDAD - ped_fac
                         ELSE CANTIDAD_DELIVERY
                    END AS 'CantTotal' ,
                    /*
exec SpCargarComanda '01','22','20140923',1
*/
                    CASE WHEN cantidad_delivery IS NULL
                         THEN ( Ped_Cantidad - ped_FAC )
                         ELSE CANTIDAD_DELIVERY - ped_fac
                    END AS 'Faltante' ,
                    CASE WHEN cantidad_delivery IS NULL
                         THEN ( Ped_Cantidad - ped_FAC ) * Ped_Precio
                         ELSE cantidad_delivery * ped_precio
                    END AS 'Importe' ,
                    --( Ped_Cantidad - ped_FAC ) * Ped_Precio AS 'Importe' ,
                    ped_numsec AS 'sec' ,
                    ped_aprobado AS 'apro' ,
                    ped_FAC AS 'aten' ,
                    ped_unidad AS 'uni' ,
                    PED_NUMSEC ,
                    PED_CTA AS 'CUENTA' ,
                    --CASE WHEN ( SELECT  COALESCE(XA.ART_CALIDAD, 1)
                    --            FROM    dbo.ARTI xa
                    --            WHERE   XA.ART_CODCIA = @CODCIA
                    --                    AND XA.ART_KEY = p.PED_CODART
                    --          ) = 0 THEN 1
                    --     ELSE 0
                    --END AS 'ICBPER' ,
                    @ICBPER AS 'GEN_ICBPER' 
                    --COALESCE(t.ICBPER*ped_cantidad, 0) AS 'COMBO_ICBPER'
					,0 AS 'COMBO_ICBPER'
					,CASE (SELECT TOP 1 COALESCE(a2.ART_BOLSAS,0) FROM dbo.ARTI a2 WHERE a2.ART_CODCIA = p.ped_codcia AND a2.ART_KEY = p.ped_codart) WHEN 0 THEN 0 ELSE
					(SELECT TOP 1 COALESCE(a2.ART_BOLSAS,0) * @ICBPER FROM dbo.ARTI a2 WHERE a2.ART_CODCIA = p.ped_codcia AND a2.ART_KEY = p.ped_codart) END AS icbper
            FROM    pedidos p
                    INNER JOIN ALLOG A ON p.PED_TRANSP = a.ALL_NUMOPER
                                          AND A.ALL_CODCIA = @CODCIA
                                          AND A.ALL_FECHA_DIA = @FECHA
                                          AND A.ALL_FLAG_EXT = 'N'
                    --LEFT JOIN @TBLICBPER t ON P.PED_CODART = t.CODART
            WHERE   ped_codclie = @CodMesa
                    AND ped_codcia = @CodCia
                    AND ped_fecha = @Fecha
                    AND ped_estado = 'N'
                    AND ped_situacion <> 'A'
                    AND ped_fac <> ped_Cantidad
                    AND ped_numfac = @tt 




        END
GO