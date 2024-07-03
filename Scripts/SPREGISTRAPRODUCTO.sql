IF EXISTS
(
    SELECT TOP 1
        s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPREGISTRAPRODUCTO'
)
BEGIN
    DROP PROC [dbo].[SPREGISTRAPRODUCTO];
END;
GO
CREATE PROC [dbo].[SPREGISTRAPRODUCTO]
    @Codcia CHAR(2) ,
    @Descrip VARCHAR(70) ,
    @alterno VARCHAR(20) ,
    @unidad VARCHAR(20) ,
    @codfam INT ,
    @codsubfam INT ,
    @stockmin NUMERIC(11, 2) ,
    @stockmax NUMERIC(11, 2) ,
    @pp1 NUMERIC(11, 4) ,
    @pp2 NUMERIC(11, 4) ,
    @pp3 NUMERIC(11, 4) ,
    @pp4 NUMERIC(11, 4) ,
    @pp5 NUMERIC(11, 4) ,
    @pp6 NUMERIC(11, 4) ,
    @pp11 NUMERIC(11, 4) ,
    @pp22 NUMERIC(11, 4) ,
    @pp33 NUMERIC(11, 4) ,
    @pp44 NUMERIC(11, 4) ,
    @pp55 NUMERIC(11, 4) ,
    @pp66 NUMERIC(11, 4) ,
    @proporcion INT ,
    @sit INT ,
    @pri INT ,
    @pri2 INT ,
    @xmlCompform VARCHAR(8000) ,
    @stock BIT = 1 ,
    @porcion BIT ,
    @preporcion MONEY,
    @flagstock CHAR(1) ,
    @COSTO MONEY,
	@BOLSAS INT =0,
	@CODBOLSA BIGINT = 0,
    @MaxCod INT OUT 
    
AS
    SET nocount ON

	IF @CODBOLSA = 0
	BEGIN
	    SET @CODBOLSA=NULL
	END

    IF @xmlCompform = ''
        BEGIN
            SET @xmlCompform = NULL
        END

    DECLARE @NroErr INT
--Bloque creacion de variables para table arti
    DECLARE @art_costo NUMERIC(4, 2) ,
        @art_margen NUMERIC(4, 2) ,
        @art_cash NUMERIC(4, 2) ,
        @art_tipo CHAR(1) ,
        @art_estado CHAR(1) ,
        @art_linea INT ,
        @art_marca INT ,
	--@art_calidad int,
        @art_plancha INT ,
        @art_unidad CHAR(5) ,
        @art_ex_igv CHAR(1) ,
        @art_decimales INT ,
        @art_por_igv NUMERIC(4, 2) ,
        @art_cospro_ant NUMERIC(4, 2) ,
        @art_cuenta_contab_c VARCHAR(30) ,
        @art_cuenta_contab VARCHAR(60) ,
        @art_moneda CHAR(1) ,
	--@art_situacion char(1),
        @art_orden INT ,
        @art_codclie NUMERIC(4, 2) ,
        @art_subgru INT ,
        @art_por1 NUMERIC(4, 2) ,
        @art_por2 NUMERIC(4, 2) ,
        @art_por3 NUMERIC(4, 2) ,
        @art_por4 NUMERIC(4, 2) ,
        @art_por5 NUMERIC(4, 2) ,
        @art_por6 NUMERIC(4, 2) ,
        @art_cuenta_contab_70 VARCHAR(50) ,
        @art_cuenta_contab_69 VARCHAR(50) ,
        @art_codart2 NUMERIC(4, 2) ,
        @art_cp CHAR(1) ,
        @art_cospro NUMERIC(4, 2)


    SET @art_costo = 0
    SET @art_margen = 0
    SET @art_cash = 0
    SET @art_tipo = 'V'
    SET @art_estado = NULL
    SET @art_linea = 0
    SET @art_marca = 0
--set @art_calidad = 1
    SET @art_plancha = 0
    SET @art_unidad = NULL
    SET @art_ex_igv = ''
    SET @art_decimales = 2
    SET @art_por_igv = 0
    SET @art_cospro_ant = 0
    SET @art_cuenta_contab_c = 0
    SET @art_cuenta_contab = 0
    SET @art_moneda = 'S'
--set @art_situacion = '1'
    SET @art_orden = 1
    SET @art_codclie = 0
    SET @art_subgru = 0
    SET @art_por1 = 0
    SET @art_por2 = 0
    SET @art_por3 = 0
    SET @art_por4 = 0
    SET @art_por5 = 0
    SET @art_por6 = 0
    SET @art_cuenta_contab_70 = NULL
    SET @art_cuenta_contab_69 = NULL
    SET @art_codart2 = 0
    SET @art_cp = ''

    SET @art_cospro = 0


--=================================================
    BEGIN TRAN
    SELECT  @MaxCod = ISNULL(MAX(art_key), 2) + 1
    FROM    arti
    WHERE   art_codcia = @codcia

    INSERT  INTO Arti
            ( art_key ,
              art_codcia ,
              art_nombre ,
              art_costo ,
              art_margen ,
              art_cash ,
              art_tipo ,
              art_estado ,
              art_numero ,
              art_linea ,
              art_marca ,
              art_calidad ,
              art_plancha ,
              art_unidad ,
              art_ex_igv ,
              art_decimales ,
              art_por_igv ,
              art_cospro_ant ,
              art_cuenta_contab_c ,
              art_cuenta_contab ,
              art_moneda ,
              art_situacion ,
              art_familia ,
              art_subfam ,
              art_orden ,
              art_codclie ,
              art_subgru ,
              art_alterno ,
              art_por1 ,
              art_por2 ,
              art_por3 ,
              art_por4 ,
              art_por5 ,
              art_por6 ,
              art_stock_min ,
              art_stock_max ,
              art_cuenta_contab_70 ,
              art_cuenta_contab_69 ,
              art_codart2 ,
              art_cp ,
              art_flag_stock ,
              art_flag_cambio ,
              art_cospro ,
              art_fecha_control ,
              ART_DESCONTARSTOCK ,
              ART_PORCION
			  ,ART_BOLSAS
			  ,ART_CODBOLSA
            )
    VALUES  ( @MaxCod ,
              @Codcia ,
              @Descrip ,
              @art_costo ,
              @art_margen ,
              @art_cash ,
              @art_tipo ,
              @art_estado ,
              @PROPORCION ,
              @art_linea ,
              @art_marca ,
              @pri ,
              @art_plancha ,
              @art_unidad ,
              @art_ex_igv ,
              @art_decimales ,
              @art_por_igv ,
              @art_cospro_ant ,
              @art_cuenta_contab_c ,
              @art_cuenta_contab ,
              @art_moneda ,
              @sit ,
              @codfam ,
              @codsubfam ,
              @art_orden ,
              @art_codclie ,
              @art_subgru ,
              @alterno ,
              @art_por1 ,
              @art_por2 ,
              @art_por3 ,
              @art_por4 ,
              @art_por5 ,
              @art_por6 ,
              @stockmin ,
              @stockmax ,
              @art_cuenta_contab_70 ,
              @art_cuenta_contab_69 ,
              @art_codart2 ,
              @art_cp ,
              @flagstock ,
              @pri2 ,
              @art_cospro ,
              GETDATE() ,
              @stock ,
              @PORCION
			  ,@BOLSAS
			  ,@CODBOLSA
            )


    SET @NroErr = @@Error
    IF @NroErr <> 0
        GOTO TratarError

--2. Grabando en tabla articulo

    INSERT  INTO articulo
            ( arm_codart ,
              arm_codcia ,
              arm_stock ,
              arm_ingresos ,
              arm_salidas ,
              arm_stock_ini ,
              arm_cospro ,
              arm_stock2 ,
              arm_costo_ult ,
              arm_fecha_ult ,
              arm_saldo_s ,
              arm_saldo_s2 ,
              arm_saldo_n ,
              arm_saldo_n2 ,
              arm_fecha_control
            )
    VALUES  ( @MaxCod ,
              @Codcia ,
              0 ,
              0 ,
              0 ,
              NULL ,
              @COSTO ,
              0 ,
              0 ,
              '01/01/1900' ,
              0 ,
              0 ,
              0 ,
              0 ,
              GETDATE()
            )

    SET @NroErr = @@Error
    IF @NroErr <> 0
        GOTO TratarError

--3. Grabando en table precios

    INSERT  INTO precios
            ( pre_codcia ,
              pre_codart ,
              pre_secuencia ,
              pre_por1 ,
              pre_por2 ,
              pre_por3 ,
              pre_por4 ,
              pre_por5 ,
              pre_por6 ,
              pre_costo ,
              pre_pre1 ,
              pre_pre2 ,
              pre_pre3 ,
              pre_pre4 ,
              pre_pre5 ,
              pre_pre6 ,
              pre_unidad ,
              pre_equiv ,
              pre_flag_unidad ,
              pre_costo_ant ,
              pre_pre11 ,
              pre_pre22 ,
              pre_pre33 ,
              pre_pre44 ,
              pre_pre55 ,
              pre_pre66 ,
              pre_peso ,
              pre_litro ,
              pre_costo_repo ,
              pre_pordes1 ,
              pre_pordes2 ,
              pre_pordes3 ,
              pre_pordes4 ,
              pre_pordes5 ,
     pre_fecha_control ,
              PRE_PORCION
            )
    VALUES  ( @CodCia ,
              @MaxCod ,
              0 ,
              NULL ,
              NULL ,
              NULL ,
              NULL ,
              NULL ,
              NULL ,
              0 , --pre_costo
              @pp1 ,
              @pp2 ,
              @pp3 ,
              @pp4 ,
              @pp5 ,
              @pp6 ,
              @unidad ,
              1 ,
              'A' ,
              0 ,
              @pp11 ,
              @pp22 ,
              @pp33 ,
              @pp44 ,
              @pp55 ,
              @pp66 ,
              0 ,
              0 ,
              0 ,
              NULL ,
              NULL ,
              NULL ,
              NULL ,
              NULL ,
              NULL ,
              @PREPORCION
            )

    SET @NroErr = @@Error
    IF @NroErr <> 0
        GOTO TratarError


--4. Grabando en tabla paquetes segun sea el caso
    DECLARE @idoc INT

--4.1. Primero la composición del producto osea es un combo
    IF @xmlCompform IS NOT NULL
        BEGIN
            DECLARE @tblComp TABLE ( idp INT, c DECIMAL(9, 3) )
            DECLARE @idp INT
            DECLARE @codalt VARCHAR(50) ,
                @xunidad VARCHAR(30) ,
                @c DECIMAL(9, 3)

            EXEC sp_xml_preparedocument @idoc OUTPUT, @xmlCompform
            INSERT  INTO @tblComp
                    SELECT  idp ,
                            c
                    FROM    OPENXML (@idoc, '/r/d',1)
	WITH (idp INT,c DECIMAL(9,3))

            DECLARE cComp CURSOR
            FOR
                SELECT  idp ,
                        c
                FROM    @tblComp	
            OPEN cComp
	
            FETCH cComp INTO @idp, @c

	

            WHILE ( @@Fetch_Status = 0 )
                BEGIN
                    SELECT  @codalt = a.art_alterno ,
                            @xunidad = p.pre_unidad
                    FROM    arti a
                            INNER JOIN precios p ON a.art_key = p.pre_codart
                                                    AND a.art_codcia = p.pre_codcia
                    WHERE   art_codcia = @codcia
                            AND art_key = @idp

                    INSERT  INTO paquetes
                            ( pa_codcia ,
                              pa_codpa ,
                              pa_codart ,
                              alterno ,
                              pa_cantidad ,
                              pa_unidad ,
                              pa_equiv ,
                              pa_flag_anulado ,
                              pa_prom ,
                              pa_fecha_ini ,
                              pa_fecha_fin
		                    )
                    VALUES  ( @codcia ,
                              @MaxCod ,
                              @idp ,
                              @codalt ,
                              0 ,
                              @xunidad ,
                              1 ,
                              'S' ,
                              @c ,
                              NULL ,
                              NULL
		                    )

                    SET @NroErr = @@Error
                    IF @NroErr <> 0
                        GOTO TratarError

                    FETCH cComp INTO @idp, @c
                END
	
            CLOSE cComp
            DEALLOCATE cComp


	
--SELECT * FROM @tblComp
--	select 'james'
        END

    COMMIT TRAN

    RETURN
    TratarError:
    ROLLBACK TRAN
GO