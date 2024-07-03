IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPMODIFICAPRODUCTO'
)
BEGIN
    DROP PROC [dbo].[SPMODIFICAPRODUCTO];
END;
GO
/*
exec SpModificaProducto '01','POLLITO','146','UNIDAD',3,19,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,N'',1,0,0,146,'M'
*/
CREATE PROC [dbo].[SPMODIFICAPRODUCTO]
    @codcia CHAR(2),
    @Descrip VARCHAR(50),
    @alterno VARCHAR(20),
    @unidad VARCHAR(20),
    @codfam INT,
    @codsubfam INT,
    @stockmin NUMERIC(11, 2),
    @stockmax NUMERIC(11, 2),
    @pp1 NUMERIC(11, 4),
    @pp2 NUMERIC(11, 4),
    @pp3 NUMERIC(11, 4),
    @pp4 NUMERIC(11, 4),
    @pp5 NUMERIC(11, 4),
    @pp6 NUMERIC(11, 4),
    @pp11 NUMERIC(9, 2),
    @pp22 NUMERIC(9, 2),
    @pp33 NUMERIC(9, 2),
    @pp44 NUMERIC(9, 2),
    @pp55 NUMERIC(9, 2),
    @pp66 NUMERIC(9, 2),
    @proporcion INT,
    @sit INT,
    @pri INT,
    @pri2 INT,
    @xmlCompform VARCHAR(8000),
    @stock BIT = 1,
    @porcion BIT,
    @preporcion MONEY,
    @codigo INT,
    @TIPO CHAR(1), --ART_FLAG_STOCK
    @COSTO MONEY,
    @BOLSAS INT = 0
	,@CODBOLSA BIGINT = 0
AS
SET NOCOUNT ON;
DECLARE @NroErr INT;
BEGIN TRAN;

IF @xmlCompform = ''
BEGIN
    SET @xmlCompform = NULL;
END;
IF @CODBOLSA = 0
	BEGIN
	    SET @CODBOLSA=NULL
	END


UPDATE ARTI
SET ART_NOMBRE = @Descrip,
    ART_FAMILIA = @codfam,
    ART_SUBFAM = @codsubfam,
    ART_ALTERNO = @alterno,
    ART_STOCK_MIN = @stockmin,
    ART_STOCK_MAX = @stockmax,
    ART_CALIDAD = @pri,
    ART_FLAG_CAMBIO = @pri2,
    ART_SITUACION = @sit,
    ART_NUMERO = @proporcion,
    ART_DESCONTARSTOCK = @stock,
    ART_FLAG_STOCK = @TIPO,
    ART_PORCION = @porcion,
    ART_BOLSAS = @BOLSAS
	,ART_CODBOLSA = @CODBOLSA
WHERE ART_KEY = @codigo
      AND ART_CODCIA = @codcia;
--select * from arti
SET @NroErr = @@Error;
IF @NroErr <> 0
    GOTO TratarError;

UPDATE ARTICULO
SET ARM_COSPRO = @COSTO
WHERE ARM_CODCIA = @codcia
      AND ARM_CODART = @codigo;
SET @NroErr = @@Error;
IF @NroErr <> 0
    GOTO TratarError;

UPDATE PRECIOS
SET PRE_PRE1 = @pp1,
    PRE_PRE2 = @pp2,
    PRE_PRE3 = @pp3,
    PRE_PRE4 = @pp4,
    PRE_PRE5 = @pp5,
    PRE_PRE6 = @pp6,
    PRE_PRE11 = @pp11,
    PRE_PRE22 = @pp22,
    PRE_PRE33 = @pp33,
    PRE_PRE44 = @pp44,
    PRE_PRE55 = @pp55,
    PRE_PRE66 = @pp66,
    PRE_UNIDAD = @unidad,
    PRE_PORCION = @preporcion
WHERE PRE_CODCIA = @codcia
      AND PRE_CODART = @codigo;

SET @NroErr = @@Error;
IF @NroErr <> 0
    GOTO TratarError;

DELETE FROM PAQUETES
WHERE PA_CODCIA = @codcia
      AND PA_CODPA = @codigo;

IF @xmlCompform IS NOT NULL
BEGIN
    DECLARE @tblComp TABLE
    (
        idp INT,
        c DECIMAL(9, 3)
    );
    DECLARE @idp INT,
            @idoc INT;
    DECLARE @codalt VARCHAR(50),
            @xunidad VARCHAR(30),
            @c DECIMAL(9, 3);

    EXEC sp_xml_preparedocument @idoc OUTPUT, @xmlCompform;
    INSERT INTO @tblComp
    SELECT idp,
           c
    FROM
        OPENXML(@idoc, '/r/d', 1)WITH (idp INT, c DECIMAL(9, 3));

    DECLARE cComp CURSOR FOR SELECT idp, c FROM @tblComp;
    OPEN cComp;

    FETCH cComp
    INTO @idp,
         @c;



    WHILE (@@Fetch_Status = 0)
    BEGIN
        SELECT @codalt = a.ART_ALTERNO,
               @xunidad = p.PRE_UNIDAD
        FROM ARTI a
            INNER JOIN PRECIOS p
                ON a.ART_KEY = p.PRE_CODART
                   AND a.ART_CODCIA = p.PRE_CODCIA
        WHERE ART_CODCIA = @codcia
              AND ART_KEY = @idp;

        INSERT INTO PAQUETES
        (
            PA_CODCIA,
            PA_CODPA,
            PA_CODART,
            ALTERNO,
            PA_CANTIDAD,
            PA_UNIDAD,
            PA_EQUIV,
            PA_FLAG_ANULADO,
            PA_PROM,
            PA_FECHA_INI,
            PA_FECHA_FIN
        )
        VALUES
        (@codcia, @codigo, @idp, @codalt, 0, @xunidad, 1, 'S', @c, NULL, NULL);

        SET @NroErr = @@Error;
        IF @NroErr <> 0
            GOTO TratarError;

        FETCH cComp
        INTO @idp,
             @c;
    END;

    CLOSE cComp;
    DEALLOCATE cComp;

END;

COMMIT TRAN;
RETURN;
TratarError:
ROLLBACK TRAN;
GO