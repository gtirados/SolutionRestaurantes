IF EXISTS
(
    SELECT TOP 1
        s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPCARGARINFOAUXPRODUCTOS'
)
BEGIN
    DROP PROC [dbo].[SPCARGARINFOAUXPRODUCTOS];
END;
GO
/*
exec SpCargarInfoAuxProductos '01',146
*/
CREATE PROCEDURE [dbo].[SPCARGARINFOAUXPRODUCTOS]
    @CodCia CHAR(2) ,
    @Codigo INT
AS 
    SET nocount ON	
declare @costo money
select top 1 @costo = ARM_COSPRO from ARTICULO where ARM_CODCIA = @CodCia and ARM_CODART = @Codigo

    SELECT  a.art_Alterno AS 'CodAlt' ,
            a.art_stock_max AS 'Maximo' ,
            a.art_stock_min AS 'Minimo' ,
            p.pre_pre1 AS 'pv1' ,
            p.pre_pre2 AS 'pv2' ,
            p.pre_pre3 AS 'pv3' ,
            p.pre_pre4 AS 'pv4' ,
            p.pre_pre5 AS 'pv5' ,
            p.pre_pre6 AS 'pv6' ,
            p.pre_pre11 AS 'pvd1' ,
            p.pre_pre22 AS 'pvd2' ,
            p.pre_pre33 AS 'pvd3' ,
            p.pre_pre44 AS 'pvd4' ,
            p.pre_pre55 AS 'pvd5' ,
            p.pre_pre66 AS 'pvd6' ,
            a.art_situacion AS 'sit' ,
            a.art_calidad AS 'pri' ,
            CASE WHEN a.ART_FLAG_CAMBIO = ''
                      OR a.ART_FLAG_CAMBIO = '1' THEN 1
                 ELSE 0
            END AS 'pri2' ,
            A.ART_NUMERO AS 'proporcion' ,
            a.art_img AS 'datoimagen' ,
            ISNULL(a.ART_DESCONTARSTOCK, 0) AS 'STOCK',
            p.PRE_UNIDAD AS 'um',
            ISNULL(a.ART_PORCION,0) AS 'porcion',
            ISNULL(p.PRE_PORCION,0) AS 'preporcion',ISNULL(@costo,0) as 'COSTO',
            a.ART_CUENTA_CONTAB as 'ctacontab'
			,COALESCE(a.ART_BOLSAS,0) AS 'bolsas'
    FROM    Arti a
            INNER JOIN precios p ON a.art_key = p.pre_codart
    WHERE   a.Art_Key = @Codigo
            AND a.art_codcia = @Codcia
            AND p.PRE_CODCIA = @Codcia


    SELECT  p.pa_codart AS 'Codigo' ,
            a.ART_NOMBRE AS 'Descripcion' ,
            pr.pre_unidad AS 'Unidad' ,
            p.pa_prom AS 'Cantidad',
            ar.ARM_COSPRO as 'CostoUnit',
            ar.ARM_COSPRO*p.pa_prom as 'Costo'
    FROM    PAQUETES p
            INNER JOIN arti a ON p.pa_codcia = a.art_codcia
                                 AND p.pa_codart = a.art_key
            INNER JOIN precios pr ON p.pa_codcia = pr.pre_codcia
                                     AND p.pa_codart = pr.pre_codart
            INNER JOIN ARTICULO ar ON p.pa_codcia = ar.ARM_CODCIA
                                 AND p.pa_codart = ar.ARM_CODART                         
    WHERE   PA_CODPA = @Codigo
            AND PA_CODCIA = @Codcia
GO