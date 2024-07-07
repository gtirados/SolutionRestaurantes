IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPFACTURARCOMANDA'
)
BEGIN
    DROP PROC [dbo].[SPFACTURARCOMANDA];
END;
GO
CREATE PROC [dbo].[SPFACTURARCOMANDA]
    @codcia CHAR(2),
    @fecha DATETIME,
    @usuario VARCHAR(20),
    @SerCom VARCHAR(3),
    @nroCom INT,
    @SerDoc VARCHAR(3),
    @NroDoc INT,           --7 (NUMERO DE DOCUMENTO DE VENTA)
    @Fbg CHAR(1),
    @XmlDet VARCHAR(4000),
    @codcli INT = 1,
    @codMozo INT,
    @totalfac MONEY,       --12(TOTAL DEL DOCUMENTO DE VENTA)
    @sec INT,
    @moneda CHAR(1),
    @diascre INT = 0,
    @farjabas TINYINT,
    @dscto MONEY,
    @CODIGODOCTO CHAR(2),
    @Xmlpag VARCHAR(4000), --19 (FORMAS DE PAGO DEL DOCUMENTO)
    @PAGACON MONEY,
    @VUELTO MONEY = 0,
    @VALORVTA MONEY,
    @VIGV MONEY,
    @GRATUITO BIT,
    @CIAPEDIDO CHAR(2),
    @ALL_ICBPER DECIMAL(8, 2),
    @SERVICIO MONEY = 0,
    @MaxNumOper INT OUT,
    @AutoNumFac INT OUT
AS
SET NOCOUNT ON;

SELECT @codMozo = pc.CODMOZO
FROM dbo.PEDIDOS_CABECERA pc
WHERE pc.CODCIA = @codcia
      AND pc.FECHA = @fecha
      AND pc.NUMSER = @SerCom
      AND pc.NUMFAC = @nroCom;

DECLARE @MaxNumFac INT;
DECLARE @tblpagos TABLE
(
    idfp INT,
    fp VARCHAR(50),
    mon CHAR(1),
    monto NUMERIC(9, 2),
    ref VARCHAR(40),
    dcre INT
);

--Variables por defecto
DECLARE @MaxNumSec INT;
DECLARE @tipomov INT,
        @ALL_CODART INT;
DECLARE @ALL_CODTRA INT,
        @ALL_FLAG_EXT CHAR(1),
        @ALL_CODCLIE INT;
DECLARE @ALL_IMPORTE_AMORT NUMERIC(9, 2);
DECLARE @ALL_IMPORTE NUMERIC(9, 2),
        @ALL_CHESER VARCHAR(3);
DECLARE @ALL_SECUENCIA INT,
        @ALL_IMPORTE_DOLL NUMERIC(9, 2);
DECLARE @ALL_PRECIO NUMERIC(9, 2),
        @ALL_CODVEN INT,
        @ALL_NUMGUIA NUMERIC(9, 2);
DECLARE @ALL_FBG CHAR(1),
        @ALL_CP CHAR(1),
        @ALL_TIPDOC CHAR(2);
DECLARE @ALL_CANTIDAD NUMERIC(9, 2),
        @ALL_CODBAN NUMERIC(9, 2);
DECLARE @ALL_AUTOCON VARCHAR(40),
        @ALL_CHENUM NUMERIC(9, 2);
DECLARE @ALL_CHESEC VARCHAR(1),
        @ALL_NUMSER CHAR(3);
DECLARE @ALL_NETO NUMERIC(9, 2),
        @ALL_BRUTO NUMERIC(9, 2);
DECLARE @ALL_IMPTO NUMERIC(9, 2);
DECLARE @ALL_DESCTO NUMERIC(9, 2);
DECLARE @ALL_MONEDA_CAJA CHAR(1),
        @ALL_MONEDA_CCM CHAR(1);
DECLARE @ALL_MONEDA_CLI CHAR(1),
        @ALL_NUMDOC BIGINT;
DECLARE @ALL_LIMCRE_ANT NUMERIC(9, 2),
        @ALL_LIMCRE_ACT NUMERIC(9, 2);
DECLARE @ALL_SIGNO_ARM INT,
        @ALL_CODTRA_EXT INT,
        @ALL_SIGNO_CCM INT;
DECLARE @ALL_SIGNO_CAR INT,
        @ALL_SIGNO_CAJA INT,
        @ALL_NUMSER_C CHAR(3);
DECLARE @ALL_NUMFAC_C NUMERIC(9, 2),
        @ALL_SERDOC INT,
        @ALL_TIPO_CAMBIO NUMERIC(9, 2);
DECLARE @ALL_FLETE NUMERIC(9, 2),
        @ALL_SUBTRA VARCHAR(20),
        @ALL_FACART CHAR(1);
DECLARE @ALL_CONCEPTO VARCHAR(40),
        @ALL_SITUACION CHAR(1),
        @ALL_FLAG_SO CHAR(1);
DECLARE @ALL_IMPG1 NUMERIC(9, 2),
        @ALL_IMPG2 NUMERIC(9, 2);
DECLARE @ALL_CTAG1 CHAR(1),
        @ALL_CTAG2 CHAR(1),
        @ALL_RUC CHAR(1);
DECLARE @ALL_CODSUNAT CHAR(2),
        @ALL_SERIE_REC INT,
        @ALL_NUM_RECIBO INT;
DECLARE @COBRA BIT;
DECLARE @hora VARCHAR(12);


--1. Preguntar si el NroDoc (NumFac) existe en el allog

IF
(
    SELECT COUNT(ALL_NUMFAC)
    FROM ALLOG
    WHERE ALL_CODCIA = @codcia
          AND ALL_FBG = @Fbg
          AND ALL_NUMSER = @SerDoc
          AND ALL_NUMFAC = @NroDoc
) = 0
BEGIN
    SET @AutoNumFac = @NroDoc;
END;
ELSE
BEGIN
    SELECT @AutoNumFac = ISNULL(MAX(ALL_NUMFAC), 0) + 1
    FROM ALLOG
    WHERE ALL_CODCIA = @codcia
          AND ALL_FBG = @Fbg
          AND ALL_NUMSER = @SerDoc;
    SET @NroDoc = @AutoNumFac;
END;

--OBTENER CODIGO DE MESA
DECLARE @CODMESA VARCHAR(10);
SELECT @CODMESA = PED_CODCLIE
FROM PEDIDOS
WHERE PED_CODCIA = @CIAPEDIDO
      AND PED_FECHA = @fecha
      AND PED_NUMSER = @SerCom
      AND PED_NUMFAC = @nroCom;


--tabla pargen
DECLARE @flagfac CHAR(1);
SELECT @flagfac = PAR_FLAG_FACTURACION
FROM PARGEN
WHERE PAR_CODCIA = @codcia;

--DECLARE @serie INT

--tipo de cambio
SELECT @ALL_TIPO_CAMBIO = ISNULL(CAL_TIPO_CAMBIO, 0)
FROM CALENDARIO
WHERE CAL_CODCIA = '00'
      AND CAL_FECHA = @fecha;
--calculando all_bruto

DECLARE @impigv NUMERIC(9, 4);
DECLARE @igv AS INT;
SELECT @igv = GEN_IGV
FROM GENERAL;



SET @impigv = (100.0 + @igv) / 100.0;

SET @ALL_IMPORTE = 0;
SET @ALL_CHESER = '0';
SET @ALL_SECUENCIA = @sec;
SET @ALL_IMPORTE_DOLL = 0;
SET @ALL_PRECIO = @totalfac;
SET @ALL_CODVEN = 0;
SET @ALL_FBG = @Fbg;



--Set @ALL_CP = @allcp
--Set @ALL_TIPDOC = @alltipdoc
SET @ALL_CANTIDAD = 0;
SET @ALL_NUMGUIA = 0;
SET @ALL_CODBAN = 0;
--Set @ALL_AUTOCON = null
SET @ALL_CHENUM = 0;
SET @ALL_CHESEC = NULL;

SET @ALL_NUMSER = CAST(@SerDoc AS CHAR(3));
--select @ALL_NUMSER as 'caleta'
--select @serie as 'caleta1'
--select @SerDoc as'caleta3'
SET @ALL_NETO = @totalfac;


--Set @ALL_BRUTO = 0

--Set @ALL_IMPTO = 0
SET @ALL_DESCTO = @dscto;
SET @ALL_MONEDA_CAJA = @moneda;
SET @ALL_MONEDA_CCM = NULL;
SET @ALL_MONEDA_CLI = NULL;

SELECT @ALL_NUMDOC = ISNULL(MAX(ALL_NUMDOC), 0)
FROM ALLOG
WHERE ALL_CODCIA = @codcia;

SET @ALL_LIMCRE_ANT = 0;
SET @ALL_LIMCRE_ACT = 0;

SET @ALL_CODTRA_EXT = 2401;
SET @ALL_SIGNO_CCM = 0;

SET @tipomov = 70;
SET @ALL_NUMSER_C = '1';
SET @ALL_SERDOC = 0;
--Set @ALL_TIPO_CAMBIO = 1
SET @ALL_FLETE = 0;

SET @ALL_FACART = NULL;
SET @ALL_SUBTRA = @ALL_AUTOCON;
SET @ALL_SITUACION = NULL;
SET @ALL_FLAG_SO = 'A';
SET @ALL_IMPG1 = 0;
SET @ALL_IMPG2 = 0;
SET @ALL_CTAG1 = NULL;
SET @ALL_CTAG2 = NULL;


SET @ALL_SERIE_REC = NULL;
SET @ALL_NUM_RECIBO = 0;
SET @ALL_RUC = NULL;
SET @tipomov = 10;
SET @ALL_CODTRA = 2401;
SET @ALL_FLAG_EXT = 'N';
SET @ALL_SIGNO_ARM = -1;
SET @ALL_CONCEPTO = 'Comanda: ' + @SerCom + '-' + RTRIM(LTRIM(CAST(@nroCom AS VARCHAR(20))));

SET @ALL_CODSUNAT = @CODIGODOCTO;

DECLARE @NroError INT;

SET @MaxNumFac = @NroDoc;

DECLARE @idocp INT,
        @idfp INT,
        @monto NUMERIC(9, 2),
        @fp VARCHAR(50),
        @mon CHAR(1),
        @ref INT,
        @dcre INT;


BEGIN TRAN;



--ALMACENANDO EN ALLOG (CABECERA) - INICIO

SELECT @MaxNumOper = ISNULL(MAX(ALL_NUMOPER), 0) + 1
FROM [dbo].[ALLOG]
WHERE ALL_CODCIA = @codcia
      AND ALL_FECHA_DIA = @fecha;


SELECT @ALL_AUTOCON = RTRIM(LTRIM(SUT_DESCRIPCION)),
       @ALL_SIGNO_CAR = SUT_SIGNO_CAR,
       @ALL_SIGNO_CAJA = SUT_SIGNO_CAJA,
       @ALL_TIPDOC = SUT_TIPDOC,
       @ALL_CP = SUT_CP
FROM SUB_TRANSA
WHERE SUT_SECUENCIA = 1
      AND SUT_CODTRA = 2401;

SET @ALL_IMPORTE_AMORT = @totalfac;
SET @ALL_BRUTO = @totalfac / @impigv;
SET @ALL_IMPTO = @totalfac - @ALL_BRUTO;

INSERT INTO [dbo].[ALLOG]
(
    ALL_CODCIA,
    ALL_FECHA_DIA,
    ALL_NUMOPER,
    ALL_CODTRA,
    ALL_FLAG_EXT,
    ALL_CODCLIE,
    ALL_CODART,
    ALL_IMPORTE_AMORT,
    ALL_IMPORTE,
    ALL_CHESER,
    ALL_SECUENCIA,
    ALL_IMPORTE_DOLL,
    ALL_CODUSU,
    ALL_PRECIO,
    ALL_CODVEN,
    ALL_FBG,
    ALL_CP,
    ALL_TIPDOC,
    ALL_CANTIDAD,
    ALL_NUMGUIA,
    ALL_CODBAN,
    ALL_AUTOCON,
    ALL_CHENUM,
    ALL_CHESEC,
    ALL_NUMSER,
    ALL_NUMFAC,
    ALL_FECHA_VCTO,
    ALL_NETO,
    ALL_BRUTO,
    ALL_GASTOS,
    ALL_IMPTO,
    ALL_DESCTO,
    ALL_MONEDA_CAJA,
    ALL_MONEDA_CCM,
    ALL_MONEDA_CLI,
    ALL_NUMDOC,
    ALL_LIMCRE_ANT,
    ALL_LIMCRE_ACT,
    ALL_NUM_OPER2,
    ALL_SIGNO_ARM,
    ALL_CODTRA_EXT,
    ALL_SIGNO_CCM,
    ALL_SIGNO_CAR,
    ALL_SIGNO_CAJA,
    ALL_TIPMOV,
    ALL_NUMSER_C,
    ALL_NUMFAC_C,
    ALL_SERDOC,
    ALL_TIPO_CAMBIO,
    ALL_FLETE,
    ALL_SUBTRA,
    ALL_HORA,
    ALL_FACART,
    ALL_CONCEPTO,
    ALL_NUMOPER2,
    ALL_FECHA_ANT,
    ALL_SITUACION,
    ALL_FECHA_SUNAT,
    ALL_FLAG_SO,
    ALL_IMPG1,
    ALL_IMPG2,
    ALL_CTAG1,
    ALL_CTAG2,
    ALL_CODSUNAT,
    ALL_FECHA_PRO,
    ALL_FECHA_CAN,
    ALL_SERIE_REC,
    ALL_NUM_RECIBO,
    ALL_RUC,
    ALL_MESA,
    all_pagacon,
    all_vuelto,
    ALL_ICBPER,
    ALL_SERVICIO
)
VALUES
(@codcia, @fecha, @MaxNumOper, @ALL_CODTRA, @ALL_FLAG_EXT, @codcli, 0, @ALL_IMPORTE_AMORT, @ALL_IMPORTE, @ALL_CHESER,
 @ALL_SECUENCIA  , @ALL_IMPORTE_DOLL, @usuario, @ALL_PRECIO, @codMozo, @ALL_FBG, @ALL_CP, @ALL_TIPDOC, @ALL_CANTIDAD, @ALL_NUMGUIA,
 @ALL_CODBAN, 'CONTADO', @ALL_CHENUM, @ALL_CHESEC, @ALL_NUMSER, @MaxNumFac, DATEADD(DAY, @dcre, @fecha), @ALL_NETO,
 @VALORVTA, @dscto, @VIGV, @ALL_DESCTO, 'S', @ALL_MONEDA_CCM, @ALL_MONEDA_CLI, @ALL_NUMDOC, @ALL_LIMCRE_ANT,
 @ALL_LIMCRE_ACT, @MaxNumFac, @ALL_SIGNO_ARM, @ALL_CODTRA_EXT, @ALL_SIGNO_CCM, @ALL_SIGNO_CAR, @ALL_SIGNO_CAJA,
 @tipomov, @SerCom, @nroCom, @ALL_SERDOC, @ALL_TIPO_CAMBIO, @ALL_FLETE, 'CONTADO', GETDATE(), @ALL_FACART,
 @ALL_CONCEPTO, @MaxNumOper, @fecha, @ALL_SITUACION, @fecha, @ALL_FLAG_SO, @ALL_IMPG1, @ALL_IMPG2, @ALL_CTAG1,
 @ALL_CTAG2, @ALL_CODSUNAT, @fecha, @fecha, @ALL_SERIE_REC, @ALL_NUM_RECIBO, @ALL_RUC, @CODMESA, @PAGACON, @VUELTO,
 @ALL_ICBPER, @SERVICIO);


SET @NroError = @@ERROR;
IF @NroError <> 0
    GOTO TratarError;

--ALMACENANDO EN ALLOG (CABECERA) - FIN

--ALMACENANDO LAS FORMAS DE PAGO DEL DOCUMENTO - INICIO

SELECT TOP 1
       @COBRA = ISNULL(u.USU_COBRA, 0)
FROM dbo.USUARIOS u
WHERE u.USU_KEY = @usuario;

IF @COBRA = 1 --PERMITE COBRAR AL USUARIO LOGUEADO - INICIO
BEGIN
    EXEC sp_xml_preparedocument @idocp OUTPUT, @Xmlpag;
    INSERT INTO @tblpagos
    SELECT idfp,
           fp,
           mon,
           monto,
           ref,
           dcre
    FROM
        OPENXML(@idocp, '/r/d', 1)
        WITH
        (
            idfp INT,
            fp VARCHAR(50),
            mon CHAR(1),
            monto NUMERIC(9, 2),
            ref INT,
            dcre INT
        );
    EXEC sp_xml_removedocument @idocp;

    DECLARE cPagos CURSOR FOR
    SELECT idfp,
           fp,
           mon,
           monto,
           ref,
           dcre
    FROM @tblpagos;

    OPEN cPagos;

    FETCH cPagos
    INTO @idfp,
         @fp,
         @mon,
         @monto,
         @ref,
         @dcre;



    WHILE (@@Fetch_Status = 0)
    BEGIN

        INSERT INTO dbo.COMPROBANTE_FORMAPAGO
        (
            CODCIA,
            TIPODOCTO,
            SERIE,
            NUMERO,
            IDFORMAPAGO,
            MONTO,
			REFERENCIA,
            DIASCREDITO,
            CU_REGISTER
        )
        VALUES
        (   @codcia, -- CODCIA - char(2)
            @Fbg,    -- TIPODOCTO - char(1)
            @SerDoc, -- SERIE - varchar(3)
            @NroDoc, -- NUMERO - bigint
            @idfp,   -- IDFORMAPAGO - int
            @monto,  -- MONTO - money
			@ref,
            @dcre,   -- DIASCREDITO - int
            @usuario -- CU_REGISTER - varchar(20)
            );


        FETCH cPagos
        INTO @idfp,
             @fp,
             @mon,
             @monto,
             @ref,
             @dcre;
    END;
END;
ELSE
BEGIN

    --OBTENIENDO LA FORMA DE PAGO AL CONTADO - INICIO
    SELECT @idfp = st.SUT_SECUENCIA,
           @dcre = SUT_SIGNO_CAR
    FROM dbo.SUB_TRANSA st
    WHERE st.SUT_CODTRA = 2401
          AND st.SUT_TIPO = 'E';
    --OBTENIENDO LA FORMA DE PAGO AL CONTADO - INICIO

	--REGISTRANDO UNICO PAGO (CONTADO)

    INSERT INTO dbo.COMPROBANTE_FORMAPAGO
    (
        CODCIA,
        TIPODOCTO,
        SERIE,
        NUMERO,
        IDFORMAPAGO,
        MONTO,
		REFERENCIA,
        DIASCREDITO,
        CU_REGISTER
    )
    VALUES
    (   @codcia,   -- CODCIA - char(2)
        @Fbg,      -- TIPODOCTO - char(1)
        @SerDoc,   -- SERIE - varchar(3)
        @NroDoc,   -- NUMERO - bigint
        @idfp,     -- IDFORMAPAGO - int
        @totalfac, -- MONTO - money
		'',
        @dcre,     -- DIASCREDITO - int
        @usuario   -- CU_REGISTER - varchar(20)
        );
END;

SET @NroError = @@ERROR;
IF @NroError <> 0
    GOTO TratarError;
--ALMACENANDO LAS FORMAS DE PAGO DEL DOCUMENTO - FIN

DECLARE @tbltmp TABLE
(
    cp INT,
    st NUMERIC(9, 2),
    pr NUMERIC(9, 2),
    un VARCHAR(20),
    sc INT,
    cam VARCHAR(50)
);



DECLARE @cp INT,
        @st NUMERIC(9, 2),
        @pr NUMERIC(9, 2),
        @un VARCHAR(20),
        @sc INT,
        @cam VARCHAR(50);
DECLARE @idoc INT,
        @sa NUMERIC(9, 2);


DECLARE @impto NUMERIC(9, 2),
        @bruto AS NUMERIC(9, 2),
        @total NUMERIC(9, 2);


--grabando en facart


EXEC sp_xml_preparedocument @idoc OUTPUT, @XmlDet;
INSERT INTO @tbltmp
SELECT cp,
       st,
       pr,
       un,
       sc,
       cam
FROM
    OPENXML(@idoc, '/r/d', 1)
    WITH
    (
        cp INT,
        st NUMERIC(9, 2),
        pr NUMERIC(9, 2),
        un VARCHAR(20),
        sc INT,
        cam VARCHAR(50)
    );
EXEC sp_xml_removedocument @idoc;


--SELECT  @total = SUM(( st * pr ))
--FROM    @tbltmp

IF @GRATUITO = 0
BEGIN
    SET @total = @totalfac + ISNULL(@dscto, 0); ---caleta
    SET @impto = ROUND(@total / 1.18, 2);
    SET @bruto = @total - @impto;
END;
ELSE
BEGIN
    SET @total = 0;
END;


--CAMPO PARA DETERMINAR SI EL PRODUCTO ESTA AFECTO AL ICBPER
DECLARE @ICBPER TINYINT,
        @VALORICBPER DECIMAL(8, 2),
        @GEN_ICBPER DECIMAL(8, 2),
        @mICBPER DECIMAL(8, 2);
SET @ICBPER = 0;
SET @GEN_ICBPER = 0;
SET @mICBPER = 0;

SELECT TOP 1
       @GEN_ICBPER = g.GEN_ICBPER
FROM dbo.GENERAL g;


DECLARE cFacArt CURSOR FOR SELECT cp, st, pr, un, sc, cam FROM @tbltmp;

OPEN cFacArt;

FETCH cFacArt
INTO @cp,
     @st,
     @pr,
     @un,
     @sc,
     @cam;

WHILE (@@Fetch_Status = 0)
BEGIN
    --AFECTO AL ICBPER
    SELECT TOP 1
           @ICBPER = COALESCE(a.ART_CALIDAD, 1)
    FROM dbo.ARTI a
    WHERE a.ART_CODCIA = @codcia
          AND a.ART_KEY = @cp;
    IF @ICBPER = 0
    BEGIN
        SET @mICBPER = (@st * @GEN_ICBPER);
    END;

    --maximo nro de secuencia
    SELECT @MaxNumSec = ISNULL(MAX(FAR_NUMSEC), 0) + 1
    FROM FACART
    WHERE FAR_TIPMOV = @tipomov
          AND FAR_NUMFAC = @MaxNumFac
          AND FAR_FBG = @Fbg;

    SELECT @hora = dbo.FnDevuelveHora(GETDATE()); --Hora actual

    --obteniendo stock actual del plato
    SELECT @sa = ARM_STOCK
    FROM ARTICULO
    WHERE ARM_CODCIA = @codcia
          AND ARM_CODART = @cp;

    DECLARE @Cliente VARCHAR(30);
    DECLARE @ruc CHAR(11);

    SELECT @Cliente = RTRIM(LTRIM(CLI_NOMBRE)),
           @ruc = CLI_RUC_ESPOSO
    FROM CLIENTES
    WHERE CLI_CODCLIE = @codcli;

    IF @Fbg = 'B'
    BEGIN
        SET @ruc = '';
    END;

    DECLARE @DESCONTARSTOCK BIT,
            @TIPO CHAR(1);

    SELECT @DESCONTARSTOCK = ISNULL(a.ART_DESCONTARSTOCK, 0),
           @TIPO = RTRIM(LTRIM(a.ART_FLAG_STOCK))
    FROM dbo.ARTI a
    WHERE a.ART_CODCIA = @codcia
          AND a.ART_KEY = @cp;

    --GRABA EL PLATO
    INSERT INTO [dbo].[FACART]
    (
        FAR_TIPMOV,
        FAR_CODCIA,
        FAR_NUMSER,
        FAR_FBG,
        FAR_NUMFAC,
        FAR_NUMSEC,
        FAR_COD_SUNAT,
        FAR_CODVEN,
        FAR_STOCK,
        FAR_CODART, --10,
        FAR_CANTIDAD,
        FAR_PRECIO,
        FAR_EQUIV,
        FAR_DESCRI,
        FAR_PESO,
        FAR_SIGNO_CAR,
        FAR_SIGNO_ARM,
        FAR_KEY_DIRCLI,
        FAR_CODCLIE,
        FAR_MONEDA, --20
        FAR_EX_IGV,
        FAR_CP,
        FAR_FECHA_COMPRA,
        FAR_ESTADO,
        FAR_ESTADO2,
        FAR_COSPRO,
        FAR_COSPRO_ANT,
        FAR_IMPTO,
        FAR_TOT_FLETE,
        FAR_FLETE,  --30
        FAR_DESCTO,
        FAR_TOT_DESCTO,
        FAR_GASTOS,
        FAR_BRUTO,
        FAR_NUMDOC,
        FAR_NUMGUIA,
        FAR_SERGUIA,
        FAR_PORDESCTO1,
        FAR_COSTEO,
        FAR_COSTEO_REAL,
        FAR_TIPO_CAMBIO,
        FAR_DIAS,
        FAR_FECHA,
        FAR_NUMSER_C,
        FAR_NUMFAC_C,
        FAR_NUMOPER,
        FAR_PRECIO_NETO,
        FAR_OTRA_CIA,
        FAR_TRANSITO,
        FAR_SUBTRA,
        FAR_JABAS,
        FAR_UNIDADES,
        FAR_MORTAL,
        FAR_NUM_PRECIO,
        FAR_ORDEN_UNIDADES,
        FAR_SUBTOTAL,
        FAR_TURNO,
        FAR_CONCEPTO,
        FAR_CODUSU,
        FAR_HORA,
        FAR_NUM_LOTE,
        FAR_PEDSER,
        FAR_PEDFAC,
        FAR_PEDSEC,
        FAR_TIPDOC,
        FAR_FECHA_PRO,
        FAR_FECHA_CAN,
        FAR_FECHA_CONTROL,
        FAR_RUC,
        FAR_FLAG_SO,
        FAR_NUMOPER2,
        FAR_OC,
        CAMBIOPRODUCTO,
        FAR_ICBPER
    )
    VALUES
    (   @tipomov, @codcia, @SerDoc, @Fbg, @MaxNumFac, @MaxNumSec, @ALL_CODSUNAT, @codMozo, @sa - @st, @cp, --10
        @st, @pr, 1, @un, 0, 0, 0, 0, @codcli, @moneda,                                                    --20
        0, @ALL_CP, @fecha, @ALL_FLAG_EXT, @ALL_FLAG_EXT, 0, @dscto, @VIGV, 0, 0,                          --30
        0, @dscto,                                                                                         --total descuento facart
        0, @VALORVTA, 0, 0, 0, 0, '', '', 1, @dcre,                                                        --FAR_DIAS
        @fecha, @SerCom, @nroCom, @MaxNumOper, 0, '', '', @ALL_SUBTRA, @farjabas, 0, 0, 0, 0, @total, 0, @ALL_CONCEPTO,
        @usuario, @hora, @ALL_SECUENCIA,                                                                   --FAR_NUM_LOTE
        0, 0, NULL, @ALL_TIPDOC, @fecha, @fecha, @fecha, @ruc, 'A', @MaxNumOper, @CODMESA, CASE
                                                                                               WHEN @cam = '' THEN
                                                                                                   NULL
                                                                                               ELSE
                                                                                                   @cam
                                                                                           END, @ALL_ICBPER);

    SET @NroError = @@ERROR;
    IF @NroError <> 0
        GOTO TratarError;

    --actualizar armstock de acuerdo a lo propocionado en el grid armstoc = armstock - cantidad
    --armsalidas = armsalidas + cantidad

    --ACTUALIZANDO TABLA PEDIDOS - cantidad facturada
    --CAMBIO POR LA OPCION DE MEDIAS PORCIONES 2014-10-07
    IF
    (
        SELECT x.CANTIDAD_DELIVERY
        FROM dbo.PEDIDOS x
        WHERE x.PED_CODCIA = @CIAPEDIDO
              AND x.PED_NUMSER = @SerCom
              AND PED_NUMFAC = @nroCom
              AND PED_NUMSEC = @sc
              AND x.PED_CODART = @cp
    ) IS NULL
    BEGIN
        UPDATE PEDIDOS
        SET PED_FAC = PED_FAC + @st
        WHERE PED_CODCIA = @CIAPEDIDO
              AND PED_FECHA = @fecha
              AND PED_NUMSER = @SerCom
              AND PED_NUMFAC = @nroCom
              AND PED_NUMSEC = @sc;
    END;
    ELSE
    BEGIN
        UPDATE PEDIDOS
        SET PED_FAC = PED_CANTIDAD --AQUI
        WHERE PED_CODCIA = @CIAPEDIDO
              AND PED_FECHA = @fecha
              AND PED_NUMSER = @SerCom
              AND PED_NUMFAC = @nroCom
              AND PED_NUMSEC = @sc;
    END;

    SET @NroError = @@ERROR;
    IF @NroError <> 0
        GOTO TratarError;
    --SELECT * FROM PEDIDOS WHERE PED_NUMFAC=2

    --	select @cp,@st,@pr,@un
    SET @mICBPER = 0;
    FETCH cFacArt
    INTO @cp,
         @st,
         @pr,
         @un,
         @sc,
         @cam;
END;

CLOSE cFacArt;
DEALLOCATE cFacArt; --TERMINA EL BUCLE DEL DETALLE DEL PEDIDO

--AQUI HAY QUE VERIFICAR SI TODA LA COMANDA HA SIDO FACTURADA PARA LIBERAR LA MESA
--select * from pedidos
DECLARE @cpc INT,
        @cpf INT;


SELECT @cpf = COUNT(PED_CODUSU)
FROM PEDIDOS
WHERE PED_CODCIA = @CIAPEDIDO
      AND PED_FECHA = @fecha
      AND PED_NUMSER = @SerCom
      AND PED_NUMFAC = @nroCom
      AND PED_CANTIDAD = PED_FAC
      AND PED_ESTADO = 'N';


SELECT @cpc = COUNT(PED_CODUSU)
FROM PEDIDOS
WHERE PED_CODCIA = @CIAPEDIDO
      AND PED_FECHA = @fecha
      AND PED_NUMSER = @SerCom
      AND PED_NUMFAC = @nroCom
      AND PED_ESTADO = 'N';



IF @cpc = @cpf
BEGIN

    --LIBERA LA MESA
    IF @COBRA = 1
    BEGIN
        UPDATE MESAS
        SET MES_ESTADO = 'L'
        WHERE MES_CODMES = @CODMESA
              AND MES_CODCIA = @CIAPEDIDO;
    END;
    ELSE
    BEGIN
        UPDATE dbo.MESAS
        SET MES_ESTADO = 'U'
        WHERE MES_CODMES = @CODMESA
              AND MES_CODCIA = @CIAPEDIDO;
    END;

    --SELECT * FROM MESAS
    SET @NroError = @@ERROR;
    IF @NroError <> 0
        GOTO TratarError;

    UPDATE dbo.PEDIDOS_CABECERA
    SET FACTURADO = 1
    WHERE CODCIA = @CIAPEDIDO
          AND NUMFAC = @nroCom
          AND NUMSER = @SerCom
          AND CONVERT(VARCHAR(8), FECHA, 112) = CONVERT(VARCHAR(8), @fecha, 112);

    SET @NroError = @@ERROR;
    IF @NroError <> 0
        GOTO TratarError;

END;

--GRABANDO EN CARTERA Y CARACU SEGUN SEA EL CASO
--SOLO SI EL VALOR DE ALLSIGNOCAJA ES 1 GRABA EN DICHAS TABLAS

IF @COBRA = 1
BEGIN
    DECLARE cCartera CURSOR FOR SELECT idfp, fp, mon, monto FROM @tblpagos;

    OPEN cCartera;

    FETCH cCartera
    INTO @idfp,
         @fp,
         @mon,
         @monto;



    WHILE (@@Fetch_Status = 0)
    BEGIN

        SELECT @ALL_AUTOCON = RTRIM(LTRIM(SUT_DESCRIPCION)),
               @ALL_SIGNO_CAR = SUT_SIGNO_CAR,
               @ALL_SIGNO_CAJA = SUT_SIGNO_CAJA,
               @ALL_TIPDOC = SUT_TIPDOC,
               @ALL_CP = SUT_CP
        FROM SUB_TRANSA
        WHERE SUT_SECUENCIA = @idfp
              AND SUT_CODTRA = 2401;



        IF @ALL_SIGNO_CAR = 1
        BEGIN
            INSERT INTO CARTERA
            (
                CAR_CP,
                CAR_CODCLIE,
                CAR_CODCIA,
                CAR_SERDOC,
                CAR_NUMDOC,
                CAR_TIPDOC,
                CAR_IMPORTE,
                CAR_FECHA_INGR,
                CAR_FECHA_VCTO,
                CAR_NUM_REN,
                CAR_CODART,
                CAR_IMP_INI,
                CAR_SITUACION,
                CAR_NUMSER,
                CAR_NUMFAC,
                CAR_PRECIO,
                CAR_CONCEPTO,
                CAR_CODTRA,
                CAR_SIGNO_CAR,
                CAR_CODVEN,
                CAR_NUMGUIA,
                CAR_FBG,
                CAR_NOMBRE_BANCO,
                CAR_NUM_CHEQUE,
                CAR_SIGNO_CAJA,
                CAR_TIPMOV,
                CAR_NUMSER_C,
                CAR_NUMFAC_C,
                CAR_FECHA_VCTO_ORIG,
                CAR_COMISION,
                CAR_CODBAN,
                CAR_COBRADOR,
                CAR_MONEDA,
                CAR_SERGUIA,
                CAR_FECHA_SUNAT,
                CAR_PLACA,
                CAR_VOUCHER,
                CAR_FLAG_SO,
                CAR_NUMOPER,
                CAR_FECHA_CONTROL,
                CAR_FECHA_ENTREGA,
                CAR_FECHA_DEVO,
                cAR_CODUNIBKO
            )

            --SELECT * FROM CARTERA
            VALUES
            (   @ALL_CP, @codcli, @codcia, @SerDoc, --@serie ,
                @NroDoc, @ALL_TIPDOC, @monto, @fecha, DATEADD(DAY, @dcre, @fecha), 4, NULL, @monto, '',
                @SerDoc,                            --@serie ,
                @MaxNumFac, 0, '', 2401, @ALL_SIGNO_CAR, @codMozo, '0', @Fbg, '', '', @ALL_SIGNO_CAJA, 10,
                0,                                  --CAR_NUMSER_C,
                0,                                  --CAR_NUMFAC_C,
                DATEADD(DAY, @dcre, @fecha),        --CAR_FECHA_VCTO_ORIG,
                0,                                  --CAR_COMISION,
                0,                                  --CAR_CODBAN,
                @codMozo,                           --CAR_COBRADOR,
                @mon,                               --CAR_MONEDA,
                0,                                  --CAR_SERGUIA,
                @fecha,                             --CAR_FECHA_SUNAT,
                NULL,                               --CAR_PLACA,
                NULL,                               --CAR_VOUCHER,
                'A',                                --CAR_FLAG_SO,
                @MaxNumOper,                        --CAR_NUMOPER,
                @fecha,                             --CAR_FECHA_CONTROL,
                NULL,                               --CAR_FECHA_ENTREGA,
                NULL,                               --CAR_FECHA_DEVO,
                NULL                                --cAR_CODUNIBKO

                );



            SET @NroError = @@ERROR;
            IF @NroError <> 0
                GOTO TratarError;


            INSERT INTO CARACU
            (
                CAA_CP,
                CAA_CODCLIE,
                CAA_CODCIA,
                CAA_TIPDOC,
                CAA_FECHA,
                CAA_NUM_OPER,
                CAA_SERDOC,
                CAA_NUMDOC,
                CAA_IMPORTE,
                CAA_SALDO,
                CAA_FECHA_VCTO,
                CAA_CONCEPTO,
                CAA_SIGNO_CAR,
                CAA_SIGNO_CCM,
                CAA_ESTADO,
                CAA_SALDO_CAR,
                CAA_TOTAL,
                CAA_NUMSER,
                CAA_NUMFAC,
                CAA_NUMGUIA,
                CAA_SIGNO_CAJA,
                CAA_TIPMOV,
                CAA_FBG,
                CAA_HORA,
                CAA_CODVEN,
                CAA_CODUSU,
                CAA_NUMSER_C,
                CAA_NUMFAC_C,
                CAA_SIGNO_CAJA_REAL,
                CAA_NUMPLAN,
                CAA_NUM_CHEQUE,
                CAA_NOTA,
                CAA_SITUACION,
                CAA_NOMBRE,
                CAA_CODBAN,
                CAA_RECIBO,
                CAA_INTVEN,
                CAA_DIASV,
                CAA_DIASA,
                CAA_TASAV,
                CAA_SERGUIA,
                CAA_FECHA_COBRO,
                CAA_FLAG_SO,
                CAA_SERIE,
                CAA_TIPO_CAMBIO,
                CAA_CODTRA,
                CAA_FECHA_CONTROL
            )
            VALUES
            (   @ALL_CP, @codcli, @codcia, @ALL_TIPDOC, @fecha, @MaxNumOper, --CAA_NUM_OPER,
                @SerDoc,                                                     --@serie , ,--CAA_SERDOC,
                @NroDoc,                                                     --CAA_NUMDOC,
                @monto,                                                      --CAA_IMPORTE,
                @monto,                                                      --CAA_SALDO,
                DATEADD(DAY, @dcre, @fecha),                                 --CAA_FECHA_VCTO,
                '',                                                          --CAA_CONCEPTO,
                @ALL_SIGNO_CAR,                                              --CAA_SIGNO_CAR,
                NULL,                                                        --CAA_SIGNO_CCM,
                'N',                                                         --CAA_ESTADO,
                @monto,                                                      --CAA_SALDO_CAR,
                @monto,                                                      --CAA_TOTAL,
                @SerDoc,                                                     --@serie , ,--CAA_NUMSER,
                @MaxNumFac,                                                  --CAA_NUMFAC,
                0,                                                           --CAA_NUMGUIA,
                4,                                                           --CAA_SIGNO_CAJA,
                10,                                                          --CAA_TIPMOV,
                @Fbg,                                                        --CAA_FBG,
                GETDATE(),                                                   --CAA_HORA,
                @codMozo,                                                    --CAA_CODVEN,
                @usuario,                                                    --CAA_CODUSU,
                0,                                                           --CAA_NUMSER_C,
                0,                                                           --CAA_NUMFAC_C,
                0,                                                           --CAA_SIGNO_CAJA_REAL,
                0,                                                           --CAA_NUMPLAN,
                '',                                                          --CAA_NUM_CHEQUE,
                NULL,                                                        --CAA_NOTA,
                '',                                                          --CAA_SITUACION,
                LEFT(@Cliente, 22),                                          --CAA_NOMBRE,
                NULL,                                                        --CAA_CODBAN,
                0,                                                           --CAA_RECIBO,
                0,                                                           --CAA_INTVEN,
                0,                                                           --CAA_DIASV,
                0,                                                           --CAA_DIASA,
                0,                                                           --CAA_TASAV,
                0,                                                           --CAA_SERGUIA,
                @fecha,                                                      --CAA_FECHA_COBRO,
                'A',                                                         --CAA_FLAG_SO,
                0,                                                           --CAA_SERIE
                @ALL_TIPO_CAMBIO,                                            --CAA_TIPO_CAMBIO
                2401,                                                        --CAA_CODTRA
                NULL                                                         --CAA_FECHA_CONTROL
                );

            SET @NroError = @@ERROR;
            IF @NroError <> 0
                GOTO TratarError;

        END;

        FETCH cCartera
        INTO @idfp,
             @fp,
             @mon,
             @monto;
    END;
END;
ELSE --NO PERMITE COBRAR
BEGIN

    SELECT @ALL_AUTOCON = RTRIM(LTRIM(SUT_DESCRIPCION)),
           @ALL_SIGNO_CAR = SUT_SIGNO_CAR,
           @ALL_SIGNO_CAJA = SUT_SIGNO_CAJA,
           @ALL_TIPDOC = SUT_TIPDOC,
           @ALL_CP = SUT_CP
    FROM SUB_TRANSA
    WHERE SUT_SECUENCIA = 4
          AND SUT_CODTRA = 2401;


    INSERT INTO CARTERA
    (
        CAR_CP,
        CAR_CODCLIE,
        CAR_CODCIA,
        CAR_SERDOC,
        CAR_NUMDOC,
        CAR_TIPDOC,
        CAR_IMPORTE,
        CAR_FECHA_INGR,
        CAR_FECHA_VCTO,
        CAR_NUM_REN,
        CAR_CODART,
        CAR_IMP_INI,
        CAR_SITUACION,
        CAR_NUMSER,
        CAR_NUMFAC,
        CAR_PRECIO,
        CAR_CONCEPTO,
        CAR_CODTRA,
        CAR_SIGNO_CAR,
        CAR_CODVEN,
        CAR_NUMGUIA,
        CAR_FBG,
        CAR_NOMBRE_BANCO,
        CAR_NUM_CHEQUE,
        CAR_SIGNO_CAJA,
        CAR_TIPMOV,
        CAR_NUMSER_C,
        CAR_NUMFAC_C,
        CAR_FECHA_VCTO_ORIG,
        CAR_COMISION,
        CAR_CODBAN,
        CAR_COBRADOR,
        CAR_MONEDA,
        CAR_SERGUIA,
        CAR_FECHA_SUNAT,
        CAR_PLACA,
        CAR_VOUCHER,
        CAR_FLAG_SO,
        CAR_NUMOPER,
        CAR_FECHA_CONTROL,
        CAR_FECHA_ENTREGA,
        CAR_FECHA_DEVO,
        cAR_CODUNIBKO
    )

    --SELECT * FROM CARTERA
    VALUES
    (   @ALL_CP, @codcli, @codcia, @SerDoc,                                                           --@serie , ,
        @NroDoc, @ALL_TIPDOC, @totalfac, @fecha, DATEADD(DAY, @dcre, @fecha), 4, NULL, @totalfac, '',
        @SerDoc,                                                                                      --@serie ,
        @MaxNumFac, 0, '', 2401, @ALL_SIGNO_CAR, @codMozo, '0', @Fbg, '', '', @ALL_SIGNO_CAJA, 10, 0, --CAR_NUMSER_C,
        0,                                                                                            --CAR_NUMFAC_C,
        DATEADD(DAY, @dcre, @fecha),                                                                  --CAR_FECHA_VCTO_ORIG,
        0,                                                                                            --CAR_COMISION,
        0,                                                                                            --CAR_CODBAN,
        @codMozo,                                                                                     --CAR_COBRADOR,
        'S',                                                                                          --CAR_MONEDA,
        0,                                                                                            --CAR_SERGUIA,
        @fecha,                                                                                       --CAR_FECHA_SUNAT,
        NULL,                                                                                         --CAR_PLACA,
        NULL,                                                                                         --CAR_VOUCHER,
        'A',                                                                                          --CAR_FLAG_SO,
        @MaxNumOper,                                                                                  --CAR_NUMOPER,
        @fecha,                                                                                       --CAR_FECHA_CONTROL,
        NULL,                                                                                         --CAR_FECHA_ENTREGA,
        NULL,                                                                                         --CAR_FECHA_DEVO,
        NULL                                                                                          --cAR_CODUNIBKO


        );



    SET @NroError = @@ERROR;
    IF @NroError <> 0
        GOTO TratarError;

    INSERT INTO CARACU
    (
        CAA_CP,
        CAA_CODCLIE,
        CAA_CODCIA,
        CAA_TIPDOC,
        CAA_FECHA,
        CAA_NUM_OPER,
        CAA_SERDOC,
        CAA_NUMDOC,
        CAA_IMPORTE,
        CAA_SALDO,
        CAA_FECHA_VCTO,
        CAA_CONCEPTO,
        CAA_SIGNO_CAR,
        CAA_SIGNO_CCM,
        CAA_ESTADO,
        CAA_SALDO_CAR,
        CAA_TOTAL,
        CAA_NUMSER,
        CAA_NUMFAC,
        CAA_NUMGUIA,
        CAA_SIGNO_CAJA,
        CAA_TIPMOV,
        CAA_FBG,
        CAA_HORA,
        CAA_CODVEN,
        CAA_CODUSU,
        CAA_NUMSER_C,
        CAA_NUMFAC_C,
        CAA_SIGNO_CAJA_REAL,
        CAA_NUMPLAN,
        CAA_NUM_CHEQUE,
        CAA_NOTA,
        CAA_SITUACION,
        CAA_NOMBRE,
        CAA_CODBAN,
        CAA_RECIBO,
        CAA_INTVEN,
        CAA_DIASV,
        CAA_DIASA,
        CAA_TASAV,
        CAA_SERGUIA,
        CAA_FECHA_COBRO,
        CAA_FLAG_SO,
        CAA_SERIE,
        CAA_TIPO_CAMBIO,
        CAA_CODTRA,
        CAA_FECHA_CONTROL
    )
    VALUES
    (   @ALL_CP, @codcli, @codcia, @ALL_TIPDOC, @fecha, @MaxNumOper, --CAA_NUM_OPER,
        @SerDoc,                                                     --@serie , ,--CAA_SERDOC,
        @NroDoc,                                                     --CAA_NUMDOC,
        @totalfac,                                                   --CAA_IMPORTE,
        @totalfac,                                                   --CAA_SALDO,
        DATEADD(DAY, @dcre, @fecha),                                 --CAA_FECHA_VCTO,
        '',                                                          --CAA_CONCEPTO,
        @ALL_SIGNO_CAR,                                              --CAA_SIGNO_CAR,
        NULL,                                                        --CAA_SIGNO_CCM,
        'N',                                                         --CAA_ESTADO,
        @totalfac,                                                   --CAA_SALDO_CAR,
        @totalfac,                                                   --CAA_TOTAL,
        @SerDoc,                                                     --@serie , ,--CAA_NUMSER,
        @MaxNumFac,                                                  --CAA_NUMFAC,
        0,                                                           --CAA_NUMGUIA,
        4,                                                           --CAA_SIGNO_CAJA,
        10,                                                          --CAA_TIPMOV,
        @Fbg,                                                        --CAA_FBG,
        GETDATE(),                                                   --CAA_HORA,
        @codMozo,                                                    --CAA_CODVEN,
        @usuario,                                                    --CAA_CODUSU,
        0,                                                           --CAA_NUMSER_C,
        0,                                                           --CAA_NUMFAC_C,
        0,                                                           --CAA_SIGNO_CAJA_REAL,
        0,                                                           --CAA_NUMPLAN,
        '',                                                          --CAA_NUM_CHEQUE,
        NULL,                                                        --CAA_NOTA,
        '',                                                          --CAA_SITUACION,
        LEFT(@Cliente, 22),                                          --CAA_NOMBRE,
        NULL,                                                        --CAA_CODBAN,
        0,                                                           --CAA_RECIBO,
        0,                                                           --CAA_INTVEN,
        0,                                                           --CAA_DIASV,
        0,                                                           --CAA_DIASA,
        0,                                                           --CAA_TASAV,
        0,                                                           --CAA_SERGUIA,
        @fecha,                                                      --CAA_FECHA_COBRO,
        'A',                                                         --CAA_FLAG_SO,
        0,                                                           --CAA_SERIE
        @ALL_TIPO_CAMBIO,                                            --CAA_TIPO_CAMBIO
        2401,                                                        --CAA_CODTRA
        NULL                                                         --CAA_FECHA_CONTROL

        );

    SET @NroError = @@ERROR;
    IF @NroError <> 0
        GOTO TratarError;
END;



COMMIT TRAN;

RETURN;

TratarError:

ROLLBACK TRAN;

RETURN;

GO