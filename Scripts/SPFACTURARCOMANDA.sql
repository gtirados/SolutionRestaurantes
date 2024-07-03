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
    @codcia CHAR(2) ,
    @fecha DATETIME ,
    @usuario VARCHAR(20) ,
    @SerCom VARCHAR(3) ,
    @nroCom INT ,
    @SerDoc VARCHAR(3) ,
    @NroDoc INT , --NUMFAC
    @Fbg CHAR(1) ,
    @XmlDet VARCHAR(4000) ,
    @codcli INT = 1 ,
    @codMozo INT ,
    @totalfac MONEY ,
    @sec INT ,
    @moneda CHAR(1) ,
    @diascre INT = 0 ,
    @farjabas TINYINT ,
    @dscto MONEY ,
    @CODIGODOCTO CHAR(2),
    @Xmlpag VARCHAR(4000) ,
    @PAGACON MONEY,
    @VUELTO MONEY = 0,
    @VALORVTA MONEY,
    @VIGV MONEY,
    @GRATUITO BIT,
    @CIAPEDIDO CHAR(2),
    @ALL_ICBPER DECIMAL(8,2),
	@SERVICIO MONEY = 0,
    @MaxNumOper INT OUT ,
    @AutoNumFac INT OUT
AS
    SET nocount ON
    
    SELECT  @codMozo = PC.CODMOZO
    FROM    dbo.PEDIDOS_CABECERA pc
    WHERE   PC.CODCIA = @codcia
            AND PC.FECHA = @fecha
            AND PC.NUMSER = @SerCom
            AND PC.NUMFAC = @nroCom
    
    DECLARE @MaxNumFac INT
    DECLARE @tblpagos TABLE
        (
          idfp INT ,
          fp VARCHAR(50) ,
          mon CHAR(1) ,
          monto NUMERIC(9, 2) ,
          ref INT,
         dcre int
        )
        
--Variables por defecto
    DECLARE @MaxNumSec INT
    DECLARE @tipomov INT ,
        @ALL_CODART INT
    DECLARE @ALL_CODTRA INT ,
        @ALL_FLAG_EXT CHAR(1) ,
        @ALL_CODCLIE INT
    DECLARE @ALL_IMPORTE_AMORT NUMERIC(9, 2)
    DECLARE @ALL_IMPORTE NUMERIC(9, 2) ,
        @ALL_CHESER VARCHAR(3)
    DECLARE @ALL_SECUENCIA INT ,
        @ALL_IMPORTE_DOLL NUMERIC(9, 2)
    DECLARE @ALL_PRECIO NUMERIC(9, 2) ,
        @ALL_CODVEN INT ,
        @ALL_NUMGUIA NUMERIC(9, 2)
    DECLARE @ALL_FBG CHAR(1) ,
        @ALL_CP CHAR(1) ,
        @ALL_TIPDOC CHAR(2)
    DECLARE @ALL_CANTIDAD NUMERIC(9, 2) ,
        @ALL_CODBAN NUMERIC(9, 2)
    DECLARE @ALL_AUTOCON VARCHAR(40) ,
        @ALL_CHENUM NUMERIC(9, 2)
    DECLARE @ALL_CHESEC VARCHAR(1) ,
        @ALL_NUMSER CHAR(3)
    DECLARE @ALL_NETO NUMERIC(9, 2) ,
        @ALL_BRUTO NUMERIC(9, 2)
    DECLARE @ALL_IMPTO NUMERIC(9, 2)
    DECLARE @ALL_DESCTO NUMERIC(9, 2)
    DECLARE @ALL_MONEDA_CAJA CHAR(1) ,
        @ALL_MONEDA_CCM CHAR(1)
    DECLARE @ALL_MONEDA_CLI CHAR(1) ,
        @ALL_NUMDOC bigint
    DECLARE @ALL_LIMCRE_ANT NUMERIC(9, 2) ,
        @ALL_LIMCRE_ACT NUMERIC(9, 2)
    DECLARE @ALL_SIGNO_ARM INT ,
        @ALL_CODTRA_EXT INT ,
        @ALL_SIGNO_CCM INT
    DECLARE @ALL_SIGNO_CAR INT ,
        @ALL_SIGNO_CAJA INT ,
        @ALL_NUMSER_C CHAR(3)
    DECLARE @ALL_NUMFAC_C NUMERIC(9, 2) ,
        @ALL_SERDOC INT ,
        @ALL_TIPO_CAMBIO NUMERIC(9, 2)
    DECLARE @ALL_FLETE NUMERIC(9, 2) ,
        @ALL_SUBTRA VARCHAR(20) ,
        @ALL_FACART CHAR(1)
    DECLARE @ALL_CONCEPTO VARCHAR(40) ,
        @ALL_SITUACION CHAR(1) ,
        @ALL_FLAG_SO CHAR(1)
    DECLARE @ALL_IMPG1 NUMERIC(9, 2) ,
        @ALL_IMPG2 NUMERIC(9, 2)
    DECLARE @ALL_CTAG1 CHAR(1) ,
        @ALL_CTAG2 CHAR(1) ,
        @ALL_RUC CHAR(1)
    DECLARE @ALL_CODSUNAT char(2) ,
        @ALL_SERIE_REC INT ,
        @ALL_NUM_RECIBO INT

    DECLARE @hora VARCHAR(12)


--1. Preguntar si el NroDoc (NumFac) existe en el allog

    IF ( SELECT COUNT(ALL_NUMFAC)
         FROM   ALLOG
         WHERE  ALL_CODCIA = @CODCIA
                AND all_fbg = @fbg
                AND all_numser = @serdoc
                AND all_numfac = @nrodoc
       ) = 0
        BEGIN
            SET @AUTONUMFAC = @NRODOC
        END
    ELSE
        BEGIN
            SELECT  @AUTONUMFAC = ISNULL(MAX(ALL_NUMFAC), 0) + 1
            FROM    ALLOG
            WHERE   ALL_CODCIA = @CodCia
                    AND all_fbg = @fbg
                    AND all_numser = @serdoc
            SET @NRODOC = @AUTONUMFAC
        END

--OBTENER CODIGO DE MESA
    DECLARE @CODMESA VARCHAR(10)
    SELECT  @CODMESA = PED_CODCLIE
    FROM    PEDIDOS
    WHERE   PED_CODCIA = @CIAPEDIDO
            AND PED_FECHA = @FECHA
          AND PED_NUMSER = @SERCOM
            AND PED_NUMFAC = @NROCOM


--tabla pargen
    DECLARE @flagfac CHAR(1)
    SELECT  @flagfac = par_flag_facturacion
    FROM    pargen
    WHERE   par_codcia = @codcia

    --DECLARE @serie INT

--tipo de cambio
    SELECT  @ALL_TIPO_CAMBIO = ISNULL(cal_tipo_cambio, 0)
    FROM    calendario
    WHERE   cal_codcia = '00'
            AND cal_Fecha = @fecha
--calculando all_bruto

    DECLARE @impigv NUMERIC(9, 4)
    DECLARE @igv AS INT
    SELECT  @igv = gen_igv
    FROM    general



    SET @impigv = ( 100.0 + @igv ) / 100.0
    
    SET @ALL_IMPORTE = 0
    SET @ALL_CHESER = '0'
    SET @ALL_SECUENCIA = @sec
    SET @ALL_IMPORTE_DOLL = 0
    SET @ALL_PRECIO = @totalfac
    SET @ALL_CODVEN = 0
    SET @ALL_FBG = @Fbg



--Set @ALL_CP = @allcp
--Set @ALL_TIPDOC = @alltipdoc
    SET @ALL_CANTIDAD = 0
    SET @ALL_NUMGUIA = 0
    SET @ALL_CODBAN = 0
--Set @ALL_AUTOCON = null
    SET @ALL_CHENUM = 0
    SET @ALL_CHESEC = NULL

    SET @ALL_NUMSER = CAST(@SerDoc AS CHAR(3))
    --select @ALL_NUMSER as 'caleta'
    --select @serie as 'caleta1'
    --select @SerDoc as'caleta3'
    SET @ALL_NETO = @totalfac


--Set @ALL_BRUTO = 0

--Set @ALL_IMPTO = 0
    SET @ALL_DESCTO = @dscto
    SET @ALL_MONEDA_CAJA = @moneda
    SET @ALL_MONEDA_CCM = NULL
    SET @ALL_MONEDA_CLI = NULL
    
    SELECT  @ALL_NUMDOC = ISNULL(MAX(all_numdoc), 0)
    FROM    allog
    WHERE   all_codcia = @codcia
    
    SET @ALL_LIMCRE_ANT = 0
    SET @ALL_LIMCRE_ACT = 0

    SET @ALL_CODTRA_EXT = 2401
    SET @ALL_SIGNO_CCM = 0

    SET @TipoMov = 70
    SET @ALL_NUMSER_C = '1'
    SET @ALL_SERDOC = 0
--Set @ALL_TIPO_CAMBIO = 1
    SET @ALL_FLETE = 0

    SET @ALL_FACART = NULL
    SET @ALL_SUBTRA = @ALL_AUTOCON
    SET @ALL_SITUACION = NULL
    SET @ALL_FLAG_SO = 'A'
    SET @ALL_IMPG1 = 0
    SET @ALL_IMPG2 = 0
    SET @ALL_CTAG1 = NULL
    SET @ALL_CTAG2 = NULL


    SET @ALL_SERIE_REC = NULL
    SET @ALL_NUM_RECIBO = 0
    SET @ALL_RUC = NULL
    SET @TipoMov = 10
    SET @ALL_CODTRA = 2401
    SET @ALL_FLAG_EXT = 'N'
    SET @ALL_SIGNO_ARM = -1
    SET @ALL_CONCEPTO = 'Comanda: ' + @Sercom + '-'
        + RTRIM(LTRIM(CAST(@nrocom AS VARCHAR(20))))

SET  @ALL_CODSUNAT = @CODIGODOCTO

    --IF @fbg = 'F'
    --    BEGIN
    --        SET @ALL_CODSUNAT = 1
    --    END
    --ELSE
    --    BEGIN
    --        IF @fbg = 'B'
    --            BEGIN
    --                SET @ALL_CODSUNAT = 3
    --            END
    --    END


    DECLARE @NroError INT

    SET @MaxNumFac = @NroDoc

    DECLARE @idocp INT ,
        @idfp INT ,
        @monto NUMERIC(9, 2) ,
        @fp VARCHAR(50) ,
        @mon CHAR(1) ,
        @ref INT,@dcre int


    BEGIN TRAN
    
    DECLARE @COBRA BIT
    SELECT TOP 1
            @COBRA = ISNULL(u.USU_COBRA, 0)
    FROM    dbo.USUARIOS u
    WHERE   u.USU_KEY = @usuario
    
    
    
   
    IF @COBRA = 1 --PERMITE COBRAR AL USUARIO LOGUEADO
        BEGIN
            EXEC sp_xml_preparedocument @idocp OUTPUT, @Xmlpag
            INSERT  INTO @tblpagos
                    SELECT  idfp ,
                            fp ,
                            mon ,
                            monto ,
                            ref, dcre
                    FROM    OPENXML (@idocp, '/r/d',1)
	WITH (idfp INT,fp VARCHAR(50),mon CHAR(1),monto NUMERIC(9,2),ref INT,dcre int)
            EXEC sp_xml_removedocument @idocp

            DECLARE cPagos CURSOR
            FOR
                SELECT  idfp ,
                        fp ,
                        mon ,
                        monto ,
                        ref,dcre
                FROM    @tblpagos

            OPEN cPagos

            FETCH cPagos INTO @idfp, @fp, @mon, @monto, @ref, @dcre



            WHILE ( @@Fetch_Status = 0 )
                BEGIN
        
                    SELECT  @ALL_AUTOCON = RTRIM(LTRIM(sut_descripcion)) ,
                            @ALL_SIGNO_CAR = sut_signo_car ,
                            @ALL_SIGNO_CAJA = sut_signo_caja ,
   @ALL_TIPDOC = sut_tipdoc ,
                            @ALL_CP = sut_cp
                    FROM    sub_transa
                    WHERE   sut_secuencia = @idfp
                            AND sut_codtra = 2401
                    --Obteniendo numero maximo de operacion all_numoper
                    SELECT  @MaxNumOper = ISNULL(MAX(ALL_NUMOPER), 0) + 1
                    FROM    [dbo].[ALLOG]
                    WHERE   ALL_CODCIA = @CodCia
                            AND ALL_FECHA_DIA = @Fecha
            
                    SET @ALL_IMPORTE_AMORT = @monto - ISNULL(@dscto, 0)
                    SET @ALL_BRUTO = @monto / @impigv
                    SET @ALL_IMPTO = @monto - @ALL_BRUTO
            
                    INSERT  INTO [dbo].[ALLOG]
                            ( ALL_CODCIA ,
                              ALL_FECHA_DIA ,
                              ALL_NUMOPER ,
                              ALL_CODTRA ,
                              ALL_FLAG_EXT ,
                              ALL_CODCLIE ,
                              ALL_CODART ,
                              ALL_IMPORTE_AMORT ,
                              ALL_IMPORTE ,
                              ALL_CHESER ,--10
                              ALL_SECUENCIA ,
                              ALL_IMPORTE_DOLL ,
                              ALL_CODUSU ,
                              ALL_PRECIO ,
                              ALL_CODVEN ,
                              ALL_FBG ,
                              ALL_CP ,
                              ALL_TIPDOC ,
                              ALL_CANTIDAD ,
                              ALL_NUMGUIA , --20
                              ALL_CODBAN ,
                              ALL_AUTOCON ,
                              ALL_CHENUM ,
                              ALL_CHESEC ,
                              ALL_NUMSER ,
                              ALL_NUMFAC ,
                              ALL_FECHA_VCTO ,
                              ALL_NETO ,
                              ALL_BRUTO ,
                              ALL_GASTOS , --30
                              ALL_IMPTO ,
                              ALL_DESCTO ,
                              ALL_MONEDA_CAJA ,
                              ALL_MONEDA_CCM ,
                              ALL_MONEDA_CLI ,
                              ALL_NUMDOC ,
                              ALL_LIMCRE_ANT ,
                              ALL_LIMCRE_ACT ,
                              ALL_NUM_OPER2 ,
                              ALL_SIGNO_ARM ,
                              ALL_CODTRA_EXT ,
                              ALL_SIGNO_CCM ,
                              ALL_SIGNO_CAR ,
                              ALL_SIGNO_CAJA ,
                              ALL_TIPMOV ,
                              ALL_NUMSER_C ,
                              ALL_NUMFAC_C ,
                              ALL_SERDOC ,
                              ALL_TIPO_CAMBIO ,
       ALL_FLETE ,
                           ALL_SUBTRA ,
                              ALL_HORA ,
                              ALL_FACART ,
                              ALL_CONCEPTO ,
                              ALL_NUMOPER2 ,
                              ALL_FECHA_ANT ,
                              ALL_SITUACION ,
                              ALL_FECHA_SUNAT ,
                              ALL_FLAG_SO ,
ALL_IMPG1 ,
                              ALL_IMPG2 ,
                              ALL_CTAG1 ,
                              ALL_CTAG2 ,
                              ALL_CODSUNAT ,
                              ALL_FECHA_PRO ,
                              ALL_FECHA_CAN ,
                              ALL_SERIE_REC ,
                              ALL_NUM_RECIBO ,
                              ALL_RUC ,
                            ALL_MESA,
    ALL_PAGACON,
                              ALL_VUELTO
                              ,ALL_ICBPER,
							  ALL_SERVICIO

	                        )
                    VALUES  ( @CodCia ,
                        @Fecha ,
                              @MaxNumOper ,
                              @ALL_CODTRA ,
                              @ALL_FLAG_EXT ,
                              @codcli ,
                              0 ,
                             -- @totalfac ,
                               --,
                              case when @totalfac = 0 then 0 else @ALL_IMPORTE_AMORT end,
                              @ALL_IMPORTE ,
                              @ALL_CHESER , --10
                              @idfp ,
                              @ALL_IMPORTE_DOLL ,
                              @Usuario ,
                              @ALL_PRECIO ,
                              @codmozo ,
                              @ALL_FBG ,
                              @ALL_CP ,
                              @ALL_TIPDOC ,
                              @ALL_CANTIDAD ,
                              @ALL_NUMGUIA , --20
                              @ALL_CODBAN ,
                              @fp ,
                              @ALL_CHENUM ,
                              @ALL_CHESEC ,
                              @ALL_NUMSER ,
                              @MaxNumFac ,
                              dateadd(day,coalesce(@dcre,0),@Fecha) ,
                              @ALL_NETO ,
                              @VALORVTA ,
                              --@ALL_BRUTO,
                              @DSCTO , --30
                              @VIGV ,
                              --@ALL_IMPTO,
                              @ALL_DESCTO ,
                              @mon ,
                              @ALL_MONEDA_CCM ,
                              @ALL_MONEDA_CLI ,
                              @ALL_NUMDOC ,
                              @ALL_LIMCRE_ANT ,
                              @ALL_LIMCRE_ACT ,
                              @MaxNumFac ,
                              @ALL_SIGNO_ARM ,
                              @ALL_CODTRA_EXT ,
                              @ALL_SIGNO_CCM ,
                              @ALL_SIGNO_CAR ,
                              @ALL_SIGNO_CAJA ,
                              @TipoMov ,
                              @SerCom ,
                              @nrocom ,
                              @ALL_SERDOC ,
                              @ALL_TIPO_CAMBIO ,
                              @ALL_FLETE ,
                              @fp ,
                              GETDATE() ,
                              @ALL_FACART ,
                              @ALL_CONCEPTO ,
                              @MaxNumOper ,
                              @Fecha ,
                              @ALL_SITUACION ,
                              @Fecha ,
                              @ALL_FLAG_SO ,
                              @ALL_IMPG1 ,
                              @ALL_IMPG2 ,
                              @ALL_CTAG1 ,
                              @ALL_CTAG2 ,
                              @ALL_CODSUNAT ,
                              @Fecha ,
                              @Fecha ,
                              @ALL_SERIE_REC ,
                              @ref ,
                              @ALL_RUC ,
                            
                    @CodMesa,
  @PAGACON,@VUELTO,@ALL_ICBPER,
  @SERVICIO

                           )
                    FETCH cPagos INTO @idfp, @fp, @mon, @monto, @ref, @dcre
                END
        END
    ELSE
        BEGIN
            SELECT  @MaxNumOper = ISNULL(MAX(ALL_NUMOPER), 0) + 1
            FROM    [dbo].[ALLOG]
            WHERE   ALL_CODCIA = @CodCia
               AND ALL_FECHA_DIA = @Fecha
                            

            SELECT  @ALL_AUTOCON = RTRIM(LTRIM(sut_descripcion)) ,
                    @ALL_SIGNO_CAR = sut_signo_car ,
                    @ALL_SIGNO_CAJA = sut_signo_caja ,
                    @ALL_TIPDOC = sut_tipdoc ,
                    @ALL_CP = sut_cp
            FROM    sub_transa
            WHERE   sut_secuencia = 4
                    AND sut_codtra = 2401
                            
            SET @ALL_IMPORTE_AMORT = @totalfac
            SET @ALL_BRUTO = @totalfac / @impigv
            SET @ALL_IMPTO = @totalfac - @ALL_BRUTO
                            
            INSERT  INTO [dbo].[ALLOG]
                    ( ALL_CODCIA ,
                      ALL_FECHA_DIA ,
                      ALL_NUMOPER ,
                      ALL_CODTRA ,
                      ALL_FLAG_EXT ,
                      ALL_CODCLIE ,
                      ALL_CODART ,
                      ALL_IMPORTE_AMORT ,
                      ALL_IMPORTE ,
                      ALL_CHESER ,
                      ALL_SECUENCIA ,
                      ALL_IMPORTE_DOLL ,
                      ALL_CODUSU ,
                      ALL_PRECIO ,
                      ALL_CODVEN ,
                      ALL_FBG ,
                      ALL_CP ,
                      ALL_TIPDOC ,
                      ALL_CANTIDAD ,
                      ALL_NUMGUIA ,
                      ALL_CODBAN ,
                      ALL_AUTOCON ,
                      ALL_CHENUM ,
                      ALL_CHESEC ,
                      ALL_NUMSER ,
                      ALL_NUMFAC ,
                      ALL_FECHA_VCTO ,
                      ALL_NETO ,
                      ALL_BRUTO ,
                      ALL_GASTOS ,
                      ALL_IMPTO ,
                      ALL_DESCTO ,
                      ALL_MONEDA_CAJA ,
                      ALL_MONEDA_CCM ,
                      ALL_MONEDA_CLI ,
                      ALL_NUMDOC ,
                      ALL_LIMCRE_ANT ,
                      ALL_LIMCRE_ACT ,
                      ALL_NUM_OPER2 ,
                      ALL_SIGNO_ARM ,
                      ALL_CODTRA_EXT ,
                      ALL_SIGNO_CCM ,
                      ALL_SIGNO_CAR ,
                      ALL_SIGNO_CAJA ,
                      ALL_TIPMOV ,
                      ALL_NUMSER_C ,
                      ALL_NUMFAC_C ,
                      ALL_SERDOC ,
                      ALL_TIPO_CAMBIO ,
                      ALL_FLETE ,
                      ALL_SUBTRA ,
                      ALL_HORA ,
                      ALL_FACART ,
                      ALL_CONCEPTO ,
                      ALL_NUMOPER2 ,
                      ALL_FECHA_ANT ,
                      ALL_SITUACION ,
                      ALL_FECHA_SUNAT ,
                      ALL_FLAG_SO ,
                      ALL_IMPG1 ,
                      ALL_IMPG2 ,
                      ALL_CTAG1 ,
                      ALL_CTAG2 ,
                      ALL_CODSUNAT ,
                      ALL_FECHA_PRO ,
                      ALL_FECHA_CAN ,
                      ALL_SERIE_REC ,
                      ALL_NUM_RECIBO ,
                      ALL_RUC ,
                      ALL_MESA,ALL_PAGACON,ALL_VUELTO,ALL_ICBPER,ALL_SERVICIO
	                )
            VALUES  ( @CodCia ,
                      @Fecha ,
                      @MaxNumOper ,
                      @ALL_CODTRA ,
                      @ALL_FLAG_EXT ,
                      @codcli ,
                      0 ,
                      @ALL_IMPORTE_AMORT ,
                      @ALL_IMPORTE ,
               @ALL_CHESER ,
            4 ,
                      @ALL_IMPORTE_DOLL ,
                      @Usuario ,
                      @ALL_PRECIO ,
                      @codmozo ,
                      @ALL_FBG ,
                      @ALL_CP ,
                      @ALL_TIPDOC ,
                      @ALL_CANTIDAD ,
                      @ALL_NUMGUIA ,
                      @ALL_CODBAN ,
                      'CREDITO' ,
                      @ALL_CHENUM ,
                @ALL_CHESEC ,
                      @ALL_NUMSER ,
                      @MaxNumFac ,
      dateadd(day,@dcre,@Fecha) ,
                      @ALL_NETO ,
                      @VALORVTA ,
                      @DSCTO ,
                      @VIGV ,
                      @ALL_DESCTO ,
                      'S' ,
                      @ALL_MONEDA_CCM ,
                      @ALL_MONEDA_CLI ,
                      @ALL_NUMDOC ,
                      @ALL_LIMCRE_ANT ,
   @ALL_LIMCRE_ACT ,
                      @MaxNumFac ,
                      @ALL_SIGNO_ARM ,
                      @ALL_CODTRA_EXT ,
                      @ALL_SIGNO_CCM ,
                      @ALL_SIGNO_CAR ,
                      @ALL_SIGNO_CAJA ,
                      @TipoMov ,
                      @SerCom ,
                      @nrocom ,
                      @ALL_SERDOC ,
                      @ALL_TIPO_CAMBIO ,
                      @ALL_FLETE ,
                      'CREDITO' ,
                      GETDATE() ,
                      @ALL_FACART ,
                      @ALL_CONCEPTO ,
                      @MaxNumOper ,
                      @Fecha ,
                      @ALL_SITUACION ,
                      @Fecha ,
                      @ALL_FLAG_SO ,
                      @ALL_IMPG1 ,
                      @ALL_IMPG2 ,
                      @ALL_CTAG1 ,
                      @ALL_CTAG2 ,
                      @ALL_CODSUNAT ,
                      @Fecha ,
                      @Fecha ,
                      @ALL_SERIE_REC ,
                      @ALL_NUM_RECIBO ,
                      @ALL_RUC ,
                      @CodMesa,@PAGACON,@VUELTO,@ALL_ICBPER
					  ,@SERVICIO

                            
                    )
                          
        END
    

    SET @NroError = @@ERROR
    IF @NroError <> 0
        GOTO TratarError


    DECLARE @tbltmp TABLE
        (
          cp INT ,
          st NUMERIC(9, 2) ,
          pr NUMERIC(9, 2) ,
          un VARCHAR(20) ,
          sc INT ,
          cam VARCHAR(50)
        )
        
       
        
    DECLARE @cp INT ,
        @st NUMERIC(9, 2) ,
        @pr NUMERIC(9, 2) ,
        @un VARCHAR(20) ,
        @sc INT ,
        @cam VARCHAR(50)
    DECLARE @idoc INT ,
        @sa NUMERIC(9, 2)


    DECLARE @impto NUMERIC(9, 2) ,
        @bruto AS NUMERIC(9, 2) ,
        @total NUMERIC(9, 2)


--grabando en facart


    EXEC sp_xml_preparedocument @idoc OUTPUT, @XmlDet
    INSERT  INTO @tbltmp
            SELECT  cp ,
                    st ,
                    pr ,
                    un ,
                    sc ,
                    cam
            FROM    OPENXML (@idoc, '/r/d',1)
	WITH (cp INT,st NUMERIC(9,2), pr NUMERIC(9,2),un VARCHAR(20),sc INT,cam VARCHAR(50))
    EXEC sp_xml_removedocument @idoc
    

    --SELECT  @total = SUM(( st * pr ))
    --FROM    @tbltmp
    
    IF @GRATUITO = 0 
    BEGIN
     SET @total = @totalfac + ISNULL(@dscto, 0) ---caleta
     SET @impto = ROUND(@total / 1.18, 2)
    SET @bruto = @total - @impto
    END
    ELSE
    BEGIN
     SET @total = 0
    END
   

    --CAMPO PARA DETERMINAR SI EL PRODUCTO ESTA AFECTO AL ICBPER
    DECLARE @ICBPER TINYINT,@VALORICBPER DECIMAL(8,2),@GEN_ICBPER DECIMAL(8,2),@mICBPER DECIMAL(8,2)
    SET @ICBPER=0
    SET @GEN_ICBPER=0
    SET @mICBPER = 0
    
    SELECT TOP 1 @GEN_ICBPER = G.GEN_ICBPER FROM dbo.GENERAL g
    

    DECLARE cFacArt CURSOR
    FOR
        SELECT  cp ,
                st ,
                pr ,
                un ,
                sc ,
                cam
        FROM    @tbltmp

    OPEN cFacArt

    FETCH cFacArt INTO @cp, @st, @pr, @un, @sc, @cam

    WHILE ( @@Fetch_Status = 0 )
        BEGIN
	--AFECTO AL ICBPER
		SELECT TOP 1 @ICBPER = COALESCE(A.ART_CALIDAD,1) FROM dbo.ARTI a WHERE A.ART_CODCIA = @CODCIA AND A.ART_KEY = @CP
		IF @ICBPER = 0
		BEGIN
			SET @mICBPER = (@ST * @GEN_ICBPER)
		END
		
	--maximo nro de secuencia
            SELECT  @Maxnumsec = ISNULL(MAX(far_numsec), 0) + 1
            FROM    facart
           WHERE   far_tipmov = @TipoMov
                    AND far_numfac = @MaxNumFac
                    AND far_fbg = @fbg

            SELECT  @hora = dbo.FnDevuelveHora(GETDATE())	--Hora actual

--obteniendo stock actual del plato
            SELECT  @sa = arm_stock
            FROM    articulo
            WHERE   arm_codcia = @codcia
                    AND arm_codart = @cp

            DECLARE @Cliente VARCHAR(30)
            DECLARE @ruc CHAR(11)

            SELECT  @Cliente = RTRIM(LTRIM(Cli_Nombre)) ,
                    @ruc = cli_ruc_esposo
            FROM    CLIEntes
            WHERE   cli_codclie = @codcli

            IF @fbg = 'B'
                BEGIN
                    SET @ruc = ''
                END

            DECLARE @DESCONTARSTOCK BIT ,
                @TIPO CHAR(1)
                
            SELECT  @DESCONTARSTOCK = ISNULL(A.ART_DESCONTARSTOCK, 0) ,
                    @TIPO = RTRIM(LTRIM(A.ART_FLAG_STOCK))
            FROM    dbo.ARTI a
            WHERE   A.ART_CODCIA = @CODCIA
                    AND A.ART_KEY = @CP
                    
			--GRABA EL PLATO
            INSERT  INTO [dbo].[FACART]
                    ( FAR_TIPMOV ,
                      FAR_CODCIA ,
                      far_numser ,
                      far_fbg ,
                      far_numfac ,
                      FAR_NUMSEC ,
                      far_cod_sunat ,
                      FAR_CODVEN ,
                      FAR_STOCk ,
                      far_codart , --10,
                      far_cantidad ,
                      FAR_PRECIO ,
                      FAR_equiv ,
                      far_descri ,
                      far_PESO ,
                      far_signo_car ,
                      far_signo_arm ,
                      far_key_dircli ,
                      far_codclie ,
                      FAR_MONEDA , --20
                      FAR_EX_IGV ,
                      FAR_cp ,
                      FAR_fecha_compra ,
                      far_estado ,
                      FAR_estado2 ,
                      FAR_COSPRO ,
                      FAR_COSPRO_ANT ,
                      far_IMPTO ,
                      FAR_TOT_FLETE ,
                      FAR_FLETE , --30
                      FAR_DESCTO ,
                      FAR_TOT_DESCTO ,
                      FAR_GASTOS ,
                      FAR_BRUTO ,
                      FAR_NUMDOC ,
                      far_numguia ,
                      far_serguia ,
                      FAR_pordescto1 ,
                      FAR_costeo ,
                      FAR_COSTEO_REAL ,
                      FAR_tipo_cambio ,
                      FAR_DIAS ,
                      FAR_fecha ,
                      FAR_NUMSER_C ,
                      FAR_NUMFAC_c ,
                      FAR_NUMOPER ,
                      far_precio_neto ,
                      far_otra_cia ,
                      far_transito ,
                      far_subtra ,
                      far_JABAS ,
                      far_UNIDADES ,
                      far_mortal ,
                      far_num_precio ,
                      FAR_ORDEN_UNIDADEs ,
                      FAR_SUBTOTAL ,
                      far_turno ,
                      far_concepto ,
                      far_codusu ,
                      FAR_HORA ,
                      FAR_NUM_LOTE ,
                      FAR_PEDSER ,
                      FAR_PEDFAC ,
                      far_pedsec ,
                      FAR_TIPDOC ,
                      far_fecha_pro ,
                      far_fecha_can ,
                      far_fecha_control ,
                      far_ruc ,
                      far_flag_so ,
                      FAR_NUMOPER2 ,
        FAR_OC ,
                      CAMBIOPRODUCTO
                      ,FAR_ICBPER
		            )
            VALUES  ( @TipoMov ,
   @CodCia ,
                      @SerDoc ,
          @fbg ,
                      @MaxNumFac ,
                      @Maxnumsec ,
                      @ALL_CODSUNAT ,
                      @codmozo ,
                      @sa - @st ,
                      @cp , --10
                      @st ,
                      @pr ,
                      1 ,
                      @un ,
                      0 ,
                      0 ,
                      --CASE WHEN @DESCONTARSTOCK = 1 THEN @ALL_SIGNO_ARM
                      --     ELSE 0
                      --END ,
                      0,
                      0 ,
                      @codcli ,
                      @moneda , --20
                      0 ,
                      @ALL_CP ,
                      @Fecha ,
                      @ALL_FLAG_EXT ,
                      @ALL_FLAG_EXT ,
                      0 ,
                      @DSCTO ,
                      @VIGV ,
                      0 ,
                      0 , --30
0 ,
                      @dscto , --total descuento facart
                      0,
                      @VALORVTA, 
                      0 ,
                      0 ,
                      0 ,
                      0 ,
                      '' ,
                      '' ,
                      1 ,
                      @dcre ,	--FAR_DIAS
                      @Fecha ,
                      @SerCom ,
                      @NroCom ,
                      @MaxNumOper ,
                      0 ,
                      '' ,
                      '' ,
                      @ALL_SUBTRA ,
                      @farjabas ,
                      0 ,
                      0 ,
                      0 ,
                      0 ,
                      @total ,
                      0 ,
                      @ALL_CONCEPTO ,
                      @Usuario ,
                      @hora ,
                      @ALL_SECUENCIA ,	--FAR_NUM_LOTE
                      0 ,
                      0 ,
                      NULL ,
                      @ALL_TIPDOC ,
                      @Fecha ,
                      @Fecha ,
                      @Fecha ,
                      @ruc ,
                      'A' ,
                      @MaxNumOper ,
                      @CodMesa ,
                      CASE WHEN @cam = '' THEN NULL
                           ELSE @CAM
                      END,
                      @ALL_ICBPER
                      
                    )

            SET @NroError = @@ERROR
            IF @NroError <> 0
                GOTO TratarError
/*
            IF @TIPO = 'C' --COMBO 
                BEGIN
                    DECLARE @TBLCOMPOSICION TABLE
                        (
                          IDPRODUCTO BIGINT ,
                          CANTIDAD NUMERIC(13, 4) ,
                          INDICE INT IDENTITY
                        )
                    DECLARE @MIN INT ,
                        @MAX INT ,
                        @IDC BIGINT ,
                        @CANTIDAD NUMERIC(13, 4)
	                
                    INSERT  INTO @TBLCOMPOSICION
                            ( IDPRODUCTO ,
                              CANTIDAD
                            )
                            SELECT  P.PA_CODART ,
                                    p.PA_PROM
                            FROM    PAQUETES P
                                    INNER JOIN ARTI A ON P.PA_CODCIA = A.ART_CODCIA
                                                         AND P.PA_CODART = A.ART_KEY
                            WHERE   P.PA_CODPA = @cp
                                    AND P.PA_CODCIA = @codcia
                                    AND A.ART_DESCONTARSTOCK = 1
					
                    SELECT  @MIN = MIN(INDICE)
                    FROM    @TBLCOMPOSICION 
                    SELECT  @MAX = MAX(INDICE)
                    FROM    @TBLCOMPOSICION 
                    
					--VARIABLE PARA EL NRO DE SECUENCIA MAXIMA
                    DECLARE @NROSEC INT

                    SELECT  @NROSEC = ISNULL(MAX(far_numsec), 0) + 1
                    FROM    facart
                    WHERE   far_tipmov = @TipoMov
                            AND far_numfac = @MaxNumFac
                            AND far_fbg = @fbg
                
                    WHILE @MIN <= @MAX
            BEGIN

                            SELECT  @IDC = IDPRODUCTO ,
                                    @CANTIDAD = CANTIDAD
      FROM    @TBLCOMPOSICION
                            WHERE   INDICE = @MIN
                            
                          /*  INSERT  INTO [dbo].[FACART]
                                    ( FAR_TIPMOV ,
                                      FAR_CODCIA ,
                                      far_numser ,
                                      far_fbg ,
                                      far_numfac ,
                                      FAR_NUMSEC ,
              far_cod_sunat ,
                                      FAR_CODVEN ,
                                      FAR_STOCk ,
                                      far_codart ,
                                      far_cantidad ,
                                      FAR_PRECIO ,
                                      FAR_equiv ,
                                      far_descri ,
                                      far_PESO ,
                                      far_signo_car ,
                                      far_signo_arm ,
                                      far_key_dircli ,
                                      far_codclie ,
                                      FAR_MONEDA ,
                                      FAR_EX_IGV ,
                                      FAR_cp ,
                                      FAR_fecha_compra ,
                                      far_estado ,
                                      FAR_estado2 ,
                                      FAR_COSPRO ,
                                      FAR_COSPRO_ANT ,
                                      far_IMPTO ,
                                      FAR_TOT_FLETE ,
                                      FAR_FLETE ,
                                      FAR_DESCTO ,
                                      FAR_TOT_DESCTO ,
                                      FAR_GASTOS ,
                                      FAR_BRUTO ,
                                      FAR_NUMDOC ,
                                      far_numguia ,
                                      far_serguia ,
                                      FAR_pordescto1 ,
                                      FAR_costeo ,
                                      FAR_COSTEO_REAL ,
                                      FAR_tipo_cambio ,
                                      FAR_DIAS ,
                                      FAR_fecha ,
                                      FAR_NUMSER_C ,
                                      FAR_NUMFAC_c ,
                                      FAR_NUMOPER ,
                                      far_precio_neto ,
                                      far_otra_cia ,
                                      far_transito ,
                                      far_subtra ,
                                      far_JABAS ,
                                      far_UNIDADES ,
                                      far_mortal ,
                                      far_num_precio ,
                                      FAR_ORDEN_UNIDADEs ,
                   FAR_SUBTOTAL ,
                                      far_turno ,
                                      far_concepto ,
                              far_codusu ,
                                      FAR_HORA ,
                                      FAR_NUM_LOTE ,
                                      FAR_PEDSER ,
                                      FAR_PEDFAC ,
                   far_pedsec ,
                                      FAR_TIPDOC ,
                                      far_fecha_pro ,
                                      far_fecha_can ,
                                      far_fecha_control ,
                                      far_ruc ,
                                      far_flag_so ,
                   FAR_NUMOPER2 ,
                                      FAR_OC ,
                                      FAR_VISIBLE ,
                                      CAMBIOPRODUCTO
		                            )
                                    SELECT  @TipoMov ,
                                            @CodCia ,
                                            @SerDoc ,
                                            @fbg ,
                                            @MaxNumFac ,
                                            @NROSEC ,-- @Maxnumsec ,
                                            @ALL_CODSUNAT ,
                                    @codmozo ,
                                            @sa - @st ,
                                            @IDC ,
                                            @st ,
                                            @pr ,
                                            1 ,
                                            @un ,
                                            0 ,
                                            0 ,
                                            @ALL_SIGNO_ARM ,
                                            0 ,
                                            @codcli ,
                                            @moneda ,
                                            0 ,
                                            @ALL_CP ,
                                            @Fecha ,
                                            @ALL_FLAG_EXT ,
                                            @ALL_FLAG_EXT ,
                                            0 ,
                                            @DSCTO ,
                                            @bruto ,
                                            0 ,
                                            0 ,
                                            0 ,
                                            @dscto , --tot descuento del facart
                                            0 ,
                                            @impto ,
                                            0 ,
                                            0 ,
                                            0 ,
                                            0 ,
                                            '' ,
                                            '' ,
                                            1 ,
                                            0 ,
                                            @Fecha ,
                                            @SerCom ,
                                            @NroCom ,
                                            @MaxNumOper ,
                                            0 ,
                                            '' ,
                                            '' ,
                                            @ALL_SUBTRA ,
                                            @farjabas ,
                                            0 ,
                                            0 ,
                                            0 ,
                                            0 ,
@total ,
                                            0 ,
                                            @ALL_CONCEPTO ,
                                            @Usuario ,
                                            @hora ,
                                   @ALL_SECUENCIA ,
                                            0 ,
                                            0 ,
                                            NULL ,
                      @ALL_TIPDOC ,
                                            @Fecha ,
                                            @Fecha ,
                                            @Fecha ,
                                            @ruc ,
                                            'A' ,
                                            @MaxNumOper ,
@CodMesa ,
                                            0 ,
                                            CASE WHEN @cam = '' THEN NULL
                                                 ELSE @CAM
                   END
                                    FROM    @TBLCOMPOSICION
                                    WHERE   INDICE = @MIN */
                                    
				--LINEAS NUEVAS
                            SET @NroError = @@ERROR
                            IF @NroError <> 0
                                GOTO TratarError
                --FIN LINEAS NUEVAS
            SET @NROSEC = @NROSEC + 1
                
                            UPDATE  ARTICULO
                            SET     ARM_STOCK = ARM_STOCK - ( @st * @CANTIDAD ) ,
                                    ARM_SALIDAS = ARM_SALIDAS + ( @st
                                                              * @CANTIDAD )
                            WHERE   ARM_CODART = @IDC
                                    AND ARM_CODCIA = @codcia
                                    --LINEAS NUEVAS
                            SET @NroError = @@ERROR
                            IF @NroError <> 0
                                GOTO TratarError
                --FIN LINEAS NUEVAS
                            SET @MIN = @MIN + 1
                        END
                END
            IF @DESCONTARSTOCK = 1
                BEGIN
                    IF @TIPO = 'P' --PLATO
                        BEGIN
                            UPDATE  ARTICULO
                            SET     ARM_STOCK = ARM_STOCK - @st ,
                                    ARM_SALIDAS = ARM_SALIDAS + @st
                            WHERE   ARM_CODART = @cp
                                    AND ARM_CODCIA = @codcia	
                        END
                END

            SET @NroError = @@ERROR
            IF @NroError <> 0
                GOTO TratarError
                
                */
                
--actualizar armstock de acuerdo a lo propocionado en el grid armstoc = armstock - cantidad
--armsalidas = armsalidas + cantidad
     
--ACTUALIZANDO TABLA PEDIDOS - cantidad facturada
--CAMBIO POR LA OPCION DE MEDIAS PORCIONES 2014-10-07
            IF ( SELECT X.CANTIDAD_DELIVERY
                 FROM   dbo.PEDIDOS x
                 WHERE  x.PED_CODCIA = @CIAPEDIDO
                        AND X.PED_NUMSER = @SERCOM
                        AND PED_NUMFAC = @NROCOM
                        AND PED_NUMSEC = @SC
                        AND X.PED_CODART = @CP
               ) IS NULL
                BEGIN
                    UPDATE  PEDIDOS
                    SET     PED_FAC = PED_FAC + @st
                    WHERE   PED_CODCIA = @CIAPEDIDO
                            AND PED_FECHA = @FECHA
                            AND PED_NUMSER = @SERCOM
                            AND PED_NUMFAC = @NROCOM
                            AND PED_NUMSEC = @sc
                END
            ELSE
                BEGIN
                    UPDATE  PEDIDOS
                    SET     PED_FAC = PED_CANTIDAD--AQUI
                    WHERE   PED_CODCIA = @CIAPEDIDO
                            AND PED_FECHA = @FECHA
                            AND PED_NUMSER = @SERCOM
                            AND PED_NUMFAC = @NROCOM
                            AND PED_NUMSEC = @sc
                END
           
            SET @NroError = @@ERROR
            IF @NroError <> 0
                GOTO TratarError
--SELECT * FROM PEDIDOS WHERE PED_NUMFAC=2

--	select @cp,@st,@pr,@un
			SET @mICBPER = 0
            FETCH cFacArt INTO @cp, @st, @pr, @un, @sc, @cam
        END

    CLOSE cFacArt
    DEALLOCATE cFacArt --TERMINA EL BUCLE DEL DETALLE DEL PEDIDO

--AQUI HAY QUE VERIFICAR SI TODA LA COMANDA HA SIDO FACTURADA PARA LIBERAR LA MESA
--select * from pedidos
    DECLARE @cpc INT ,
        @cpf INT

 
    SELECT  @cpf = COUNT(ped_codusu)
    FROM    pedidos
    WHERE   PED_CODCIA = @CIAPEDIDO
            AND PED_FECHA = @FECHA
            AND PED_NUMSER = @SERCOM
            AND PED_NUMFAC = @NROCOM
            AND ped_cantidad = ped_Fac
AND ped_estado = 'N'


    SELECT  @cpc = COUNT(ped_codusu)
    FROM    pedidos
    WHERE   PED_CODCIA = @CIAPEDIDO
            AND PED_FECHA = @FECHA
            AND PED_NUMSER = @SERCOM
            AND PED_NUMFAC = @NROCOM
            AND ped_estado = 'N'



    IF @cpc = @cpf
        BEGIN

	--LIBERA LA MESA
            IF @COBRA = 1
                BEGIN
					UPDATE  MESAS
					SET     MES_ESTADO = 'L'
					WHERE   MES_CODMES = @CODMESA
					AND MES_CODCIA = @CIAPEDIDO
                END
            ELSE
                BEGIN
                    UPDATE  dbo.MESAS
                    SET     MES_ESTADO = 'U'
                    WHERE   MES_CODMES = @CODMESA
                            AND MES_CODCIA = @CIAPEDIDO
                END
            
--SELECT * FROM MESAS
            SET @NroError = @@ERROR
            IF @NroError <> 0
                GOTO TratarError
                
            UPDATE  dbo.PEDIDOS_CABECERA
            SET     FACTURADO = 1
            WHERE   CODCIA = @CIAPEDIDO
                    AND NUMFAC = @nroCom
                    AND NUMSER = @SerCom
                    AND CONVERT(VARCHAR(8), FECHA, 112) = CONVERT(VARCHAR(8), @FECHA, 112)
                    
            SET @NroError = @@ERROR
            IF @NroError <> 0
                GOTO TratarError
                
        END

--GRABANDO EN CARTERA Y CARACU SEGUN SEA EL CASO
--SOLO SI EL VALOR DE ALLSIGNOCAJA ES 1 GRABA EN DICHAS TABLAS

    IF @COBRA = 1
        BEGIN
            DECLARE cCartera CURSOR
            FOR
                SELECT  idfp ,
                        fp ,
                        mon ,
                        monto
                FROM    @tblpagos

            OPEN cCartera

            FETCH cCartera INTO @idfp, @fp, @mon, @monto



            WHILE ( @@Fetch_Status = 0 )
                BEGIN
        
                    SELECT  @ALL_AUTOCON = RTRIM(LTRIM(sut_descripcion)) ,
                            @ALL_SIGNO_CAR = sut_signo_car ,
                            @ALL_SIGNO_CAJA = sut_signo_caja ,
                            @ALL_TIPDOC = sut_tipdoc ,
                            @ALL_CP = sut_cp
                    FROM    sub_transa
                    WHERE   sut_secuencia = @idfp
                            AND sut_codtra = 2401
                    
                    
                    
                    IF @ALL_SIGNO_CAR = 1
                        BEGIN
                            INSERT  INTO CARTERA
                                    ( CAR_CP ,
                                      CAR_CODCLIE ,
                                      CAR_CODCIA ,
                                      CAR_SERDOC ,
                                      CAR_NUMDOC ,
                                      CAR_TIPDOC ,
      CAR_IMPORTE ,
                                      CAR_FECHA_INGR ,
                                      CAR_FECHA_VCTO ,
                                      CAR_NUM_REN ,
                                      CAR_CODART ,
                                      CAR_IMP_INI ,
                                      CAR_SITUACION ,
                            CAR_NUMSER ,
                                      CAR_NUMFAC ,
                                      CAR_PRECIO ,
                                      CAR_CONCEPTO ,
                                      CAR_CODTRA ,
                                      CAR_SIGNO_CAR ,
                                      CAR_CODVEN ,
                                      CAR_NUMGUIA ,
                                      CAR_FBG ,
                                      CAR_NOMBRE_BANCO ,
                                     CAR_NUM_CHEQUE ,
                                      CAR_SIGNO_CAJA ,
                                      CAR_TIPMOV ,
                                      CAR_NUMSER_C ,
                                 CAR_NUMFAC_C ,
                                      CAR_FECHA_VCTO_ORIG ,
                                      CAR_COMISION ,
                                      CAR_CODBAN ,
                                      CAR_COBRADOR ,
                                      CAR_MONEDA ,
                                      CAR_SERGUIA ,
                                     CAR_FECHA_SUNAT ,
                                      CAR_PLACA ,
                                      CAR_VOUCHER ,
                                      CAR_FLAG_SO ,
                                      CAR_NUMOPER ,
                                      CAR_FECHA_CONTROL ,
                                      CAR_FECHA_ENTREGA ,
                                      CAR_FECHA_DEVO ,
                                      cAR_CODUNIBKO
                                    )

--SELECT * FROM CARTERA
                            VALUES  ( @ALL_CP ,
                                      @codcli ,
                                      @codcia ,
                                      @serdoc ,--@serie ,
                                      @NroDoc ,
                                      @ALL_TIPDOC ,
                                      @monto ,
                                      @fecha ,
                                      DATEADD(day, @dcre, @fecha) ,
                                      4 ,
                                      NULL ,
                                      @monto ,
                                      '' ,
                                         @serdoc ,--@serie ,
                                      @MaxNumFac ,
                                      0 ,
                                      '' ,
                                      2401 ,
                                      @ALL_SIGNO_CAR ,
                                      @codMozo ,
                                      '0' ,
                                      @Fbg ,
                                      '' ,
                                      '' ,
                                      @ALL_SIGNO_CAJA ,
                                      10 ,
                                      0 ,--CAR_NUMSER_C,
                                      0 ,--CAR_NUMFAC_C,
                                      DATEADD(day, @dcre, @fecha) ,--CAR_FECHA_VCTO_ORIG,
                                      0 ,--CAR_COMISION,
                                      0 ,--CAR_CODBAN,
                                      @codMozo ,--CAR_COBRADOR,
                                      @mon ,--CAR_MONEDA,
                                      0 ,--CAR_SERGUIA,
                                      @fecha ,--CAR_FECHA_SUNAT,
                                      NULL ,--CAR_PLACA,
                                      NULL ,--CAR_VOUCHER,
                         'A' ,--CAR_FLAG_SO,
                                      @MaxNumOper ,--CAR_NUMOPER,
                                      @fecha ,--CAR_FECHA_CONTROL,
                                      NULL ,--CAR_FECHA_ENTREGA,
                                      NULL ,--CAR_FECHA_DEVO,
                                      NULL--cAR_CODUNIBKO

                                    )



                            SET @NroError = @@ERROR
                            IF @NroError <> 0
                                GOTO TratarError


                            INSERT  INTO caracu
                                    ( CAA_CP ,
                                      CAA_CODCLIE ,
                                      CAA_CODCIA ,
                                      CAA_TIPDOC ,
                                      CAA_FECHA ,
                                      CAA_NUM_OPER ,
                                      CAA_SERDOC ,
                                      CAA_NUMDOC ,
                                      CAA_IMPORTE ,
                                     CAA_SALDO ,
                                      CAA_FECHA_VCTO ,
                                      CAA_CONCEPTO ,
                                      CAA_SIGNO_CAR ,
                                      CAA_SIGNO_CCM ,
                                      CAA_ESTADO ,
                                      CAA_SALDO_CAR ,
                                      CAA_TOTAL ,
                                      CAA_NUMSER ,
                                      CAA_NUMFAC ,
                                      CAA_NUMGUIA ,
                                      CAA_SIGNO_CAJA ,
                                      CAA_TIPMOV ,
                                      CAA_FBG ,
                                      CAA_HORA ,
                                      CAA_CODVEN ,
                                      CAA_CODUSU ,
                                      CAA_NUMSER_C ,
                                      CAA_NUMFAC_C ,
                                      CAA_SIGNO_CAJA_REAL ,
                                      CAA_NUMPLAN ,
                                      CAA_NUM_CHEQUE ,
                                      CAA_NOTA ,
                                      CAA_SITUACION ,
                                      CAA_NOMBRE ,
                                      CAA_CODBAN ,
                                      CAA_RECIBO ,
                                      CAA_INTVEN ,
                                      CAA_DIASV ,
                                      CAA_DIASA ,
                                      CAA_TASAV ,
                                      CAA_SERGUIA ,
                                      CAA_FECHA_COBRO ,
                                      CAA_FLAG_SO ,
                                      CAA_SERIE ,
                                      CAA_TIPO_CAMBIO ,
                                      CAA_CODTRA ,
                                      CAA_FECHA_CONTROL
                                    )
                            VALUES  ( @ALL_CP ,
                                      @codcli ,
                                      @CODCIA ,
                                      @ALL_TIPDOC ,
                                      @fecha ,
                                      @MaxNumOper ,--CAA_NUM_OPER,
                                         @serdoc ,--@serie , ,--CAA_SERDOC,
                                      @NroDoc ,--CAA_NUMDOC,
                                      @monto ,--CAA_IMPORTE,
                                      @monto ,--CAA_SALDO,
                                      DATEADD(day, @dcre, @fecha) ,--CAA_FECHA_VCTO,
                                      '' ,--CAA_CONCEPTO,
                                      @ALL_SIGNO_CAR ,--CAA_SIGNO_CAR,
                                      NULL ,--CAA_SIGNO_CCM,
                           'N' ,--CAA_ESTADO,
                                      @monto ,--CAA_SALDO_CAR,
                                      @monto ,--CAA_TOTAL,
                                          @serdoc ,--@serie , ,--CAA_NUMSER,
                                      @MaxNumFac ,--CAA_NUMFAC,
     0 ,--CAA_NUMGUIA,
                                      4 ,--CAA_SIGNO_CAJA,
                                      10 ,--CAA_TIPMOV,
                                      @fbg ,--CAA_FBG,
                                      GETDATE() ,--CAA_HORA,
                                      @codmozo ,--CAA_CODVEN,
                                      @Usuario ,--CAA_CODUSU,
                                      0 ,--CAA_NUMSER_C,
                                      0 ,--CAA_NUMFAC_C,
                          0 ,--CAA_SIGNO_CAJA_REAL,
                                      0 ,--CAA_NUMPLAN,
                                      '' ,--CAA_NUM_CHEQUE,
                                      NULL ,--CAA_NOTA,
                                '' ,--CAA_SITUACION,
                                      LEFT(@Cliente, 22) ,--CAA_NOMBRE,
                                      NULL ,--CAA_CODBAN,
                                      0 ,--CAA_RECIBO,
                                      0 ,--CAA_INTVEN,
                                      0 ,--CAA_DIASV,
                            0 ,--CAA_DIASA,
                                      0 ,--CAA_TASAV,
                                      0 ,--CAA_SERGUIA,
                                      @fecha ,--CAA_FECHA_COBRO,
                                      'A' ,--CAA_FLAG_SO,
                                      0 ,--CAA_SERIE
                                      @ALL_TIPO_CAMBIO ,--CAA_TIPO_CAMBIO
                                      2401 ,--CAA_CODTRA
                                      NULL--CAA_FECHA_CONTROL
                                    )

                            SET @NroError = @@ERROR
                            IF @NroError <> 0
                                GOTO TratarError
                
                
                
                
                        END
            
                    FETCH cCartera INTO @idfp, @fp, @mon, @monto
                END
        END
    ELSE --NO PERMITE COBRAR
        BEGIN

            SELECT  @ALL_AUTOCON = RTRIM(LTRIM(sut_descripcion)) ,
                    @ALL_SIGNO_CAR = sut_signo_car ,
                    @ALL_SIGNO_CAJA = sut_signo_caja ,
                    @ALL_TIPDOC = sut_tipdoc ,
                    @ALL_CP = sut_cp
            FROM    sub_transa
            WHERE   sut_secuencia = 4
                    AND sut_codtra = 2401
                    
                    
            INSERT  INTO CARTERA
                    ( CAR_CP ,
                      CAR_CODCLIE ,
                      CAR_CODCIA ,
                      CAR_SERDOC ,
                      CAR_NUMDOC ,
                      CAR_TIPDOC ,
                      CAR_IMPORTE ,
                      CAR_FECHA_INGR ,
                      CAR_FECHA_VCTO ,
                      CAR_NUM_REN ,
                      CAR_CODART ,
                      CAR_IMP_INI ,
                      CAR_SITUACION ,
                      CAR_NUMSER ,
                      CAR_NUMFAC ,
                      CAR_PRECIO ,
                      CAR_CONCEPTO ,
                      CAR_CODTRA ,
                      CAR_SIGNO_CAR ,
                      CAR_CODVEN ,
                      CAR_NUMGUIA ,
                      CAR_FBG ,
                      CAR_NOMBRE_BANCO ,
                      CAR_NUM_CHEQUE ,
                      CAR_SIGNO_CAJA ,
                      CAR_TIPMOV ,
                      CAR_NUMSER_C ,
                      CAR_NUMFAC_C ,
                      CAR_FECHA_VCTO_ORIG ,
                      CAR_COMISION ,
                      CAR_CODBAN ,
                      CAR_COBRADOR ,
                      CAR_MONEDA ,
                      CAR_SERGUIA ,
                      CAR_FECHA_SUNAT ,
                      CAR_PLACA ,
                      CAR_VOUCHER ,
                      CAR_FLAG_SO ,
                      CAR_NUMOPER ,
  CAR_FECHA_CONTROL ,
                      CAR_FECHA_ENTREGA ,
                      CAR_FECHA_DEVO ,
                      cAR_CODUNIBKO
                            
                    )

--SELECT * FROM CARTERA
            VALUES  ( @ALL_CP ,
                      @codcli ,
                      @codcia ,
                          @serdoc ,--@serie , ,
                      @NroDoc ,
                      @ALL_TIPDOC ,
                      @totalfac ,
   @fecha ,
                      DATEADD(day, @dcre, @fecha) ,
                      4 ,
                      NULL ,
                      @totalfac ,
                      '' ,
                          @serdoc ,--@serie ,
     @MaxNumFac ,
                      0 ,
                      '' ,
                      2401 ,
                      @ALL_SIGNO_CAR ,
                      @codMozo ,
                      '0' ,
                      @Fbg ,
                      '' ,
                      '' ,
                      @ALL_SIGNO_CAJA ,
                      10 ,
                      0 ,--CAR_NUMSER_C,
                      0 ,--CAR_NUMFAC_C,
                      DATEADD(day, @dcre, @fecha) ,--CAR_FECHA_VCTO_ORIG,
                      0 ,--CAR_COMISION,
                      0 ,--CAR_CODBAN,
                      @codMozo ,--CAR_COBRADOR,
                      'S' ,--CAR_MONEDA,
                      0 ,--CAR_SERGUIA,
                      @fecha ,--CAR_FECHA_SUNAT,
                      NULL ,--CAR_PLACA,
                      NULL ,--CAR_VOUCHER,
                      'A' ,--CAR_FLAG_SO,
                      @MaxNumOper ,--CAR_NUMOPER,
                      @fecha ,--CAR_FECHA_CONTROL,
                      NULL ,--CAR_FECHA_ENTREGA,
                      NULL ,--CAR_FECHA_DEVO,
                      NULL--cAR_CODUNIBKO

                            
                    )



            SET @NroError = @@ERROR
            IF @NroError <> 0
                GOTO TratarError
                        
            INSERT  INTO caracu
                    ( CAA_CP ,
                      CAA_CODCLIE ,
                      CAA_CODCIA ,
                      CAA_TIPDOC ,
                      CAA_FECHA ,
                      CAA_NUM_OPER ,
                      CAA_SERDOC ,
                      CAA_NUMDOC ,
                      CAA_IMPORTE ,
                      CAA_SALDO ,
                      CAA_FECHA_VCTO ,
                      CAA_CONCEPTO ,
                      CAA_SIGNO_CAR ,
                      CAA_SIGNO_CCM ,
                      CAA_ESTADO ,
                      CAA_SALDO_CAR ,
                      CAA_TOTAL ,
                      CAA_NUMSER ,
                      CAA_NUMFAC ,
                      CAA_NUMGUIA ,
                      CAA_SIGNO_CAJA ,
                      CAA_TIPMOV ,
                      CAA_FBG ,
                      CAA_HORA ,
                      CAA_CODVEN ,
                      CAA_CODUSU ,
                      CAA_NUMSER_C ,
                      CAA_NUMFAC_C ,
                      CAA_SIGNO_CAJA_REAL ,
                      CAA_NUMPLAN ,
                      CAA_NUM_CHEQUE ,
                      CAA_NOTA ,
                      CAA_SITUACION ,
                      CAA_NOMBRE ,
                      CAA_CODBAN ,
                      CAA_RECIBO ,
                      CAA_INTVEN ,
                      CAA_DIASV ,
                      CAA_DIASA ,
                      CAA_TASAV ,
                      CAA_SERGUIA ,
                      CAA_FECHA_COBRO ,
                      CAA_FLAG_SO ,
                      CAA_SERIE ,
                      CAA_TIPO_CAMBIO ,
                      CAA_CODTRA ,
                      CAA_FECHA_CONTROL
                            
                    )
            VALUES  ( @ALL_CP ,
                      @codcli ,
                      @CODCIA ,
    @ALL_TIPDOC ,
                      @fecha ,
                      @MaxNumOper ,--CAA_NUM_OPER,
                          @serdoc ,--@serie , ,--CAA_SERDOC,
                      @NroDoc ,--CAA_NUMDOC,
                      @totalfac ,--CAA_IMPORTE,
                      @totalfac ,--CAA_SALDO,
                      DATEADD(day, @dcre, @fecha) ,--CAA_FECHA_VCTO,
                      '' ,--CAA_CONCEPTO,
                      @ALL_SIGNO_CAR ,--CAA_SIGNO_CAR,
                      NULL ,--CAA_SIGNO_CCM,
                      'N' ,--CAA_ESTADO,
                      @totalfac ,--CAA_SALDO_CAR,
                      @totalfac ,--CAA_TOTAL,
                          @serdoc ,--@serie , ,--CAA_NUMSER,
                      @MaxNumFac ,--CAA_NUMFAC,
                      0 ,--CAA_NUMGUIA,
                      4 ,--CAA_SIGNO_CAJA,
                      10 ,--CAA_TIPMOV,
                      @fbg ,--CAA_FBG,
                      GETDATE() ,--CAA_HORA,
                      @codmozo ,--CAA_CODVEN,
           @Usuario ,--CAA_CODUSU,
                      0 ,--CAA_NUMSER_C,
                      0 ,--CAA_NUMFAC_C,
                      0 ,--CAA_SIGNO_CAJA_REAL,
                      0 ,--CAA_NUMPLAN,
                      '' ,--CAA_NUM_CHEQUE,
                      NULL ,--CAA_NOTA,
                      '' ,--CAA_SITUACION,
                      LEFT(@Cliente, 22) ,--CAA_NOMBRE,
                      NULL ,--CAA_CODBAN,
                      0 ,--CAA_RECIBO,
                      0 ,--CAA_INTVEN,
                      0 ,--CAA_DIASV,
                      0 ,--CAA_DIASA,
                      0 ,--CAA_TASAV,
                      0 ,--CAA_SERGUIA,
                      @fecha ,--CAA_FECHA_COBRO,
                      'A' ,--CAA_FLAG_SO,
                      0 ,--CAA_SERIE
                      @ALL_TIPO_CAMBIO ,--CAA_TIPO_CAMBIO
                      2401 ,--CAA_CODTRA
                      NULL--CAA_FECHA_CONTROL
                            
                    )

            SET @NroError = @@ERROR
            IF @NroError <> 0
                GOTO TratarError
        END

   

    COMMIT TRAN
/*end

else

begin
	raiserror('Nro de Documento ya Emitido.',16,1)
end
*/
    RETURN

    TratarError:

    ROLLBACK TRAN

    RETURN

GO