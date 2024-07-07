IF EXISTS
(
    SELECT TOP 1
           s.SPECIFIC_NAME
    FROM INFORMATION_SCHEMA.ROUTINES s
    WHERE s.ROUTINE_TYPE = 'PROCEDURE'
          AND s.ROUTINE_NAME = 'SPPRINTCOMANDA2'
)
BEGIN
    DROP PROC [dbo].[SPPRINTCOMANDA2];
END;
GO
/*
precuenta
exec SpPrintComanda2 '01','100',1,'34,','0,',1,' '
go
imprimir comanda
exec SpPrintComanda2 '01','100',1,'34,','0,'


exec SpPrintComanda2 '01','100',1,'34,113,109,','0,1,2,',1,' '
*/

CREATE PROCEDURE [dbo].[SPPRINTCOMANDA2]
    @CodCia CHAR(2),
    @NumSer CHAR(3),
    @NumFac INT,
    @xdet VARCHAR(4000) = NULL,
    @xnumsec VARCHAR(4000) = NULL,
    @precuenta BIT = NULL,
    @CTA CHAR(1) = NULL
--With Encryption
AS
SET NOCOUNT ON;
DECLARE @tbltmp TABLE
(
    cp INT
);
DECLARE @idoc INT;

--productos en general
DECLARE @tbldata TABLE
(
    PED_FECHA DATETIME,
    NROCOMANDA VARCHAR(10),
    PED_CANTIDAD MONEY,
    PED_PRECIO MONEY,
    PED_IGV MONEY,
    PED_BRUTO MONEY,
    PED_HORA VARCHAR(15),
    PED_MONEDA CHAR(1),
    PED_SUBTOTAL MONEY,
    ART_NOMBRE VARCHAR(80),
    CLI_NOMBRE VARCHAR(80),
    VEM_NOMBRE VARCHAR(60),
    PED_OFERTA VARCHAR(300),
    PED_CLIENTE VARCHAR(120),
    ped_familia INT,
    codprod BIGINT,
    flag CHAR(1),
    actual DATETIME,
    FAMILIA VARCHAR(100),
    CARACTERISTICAS VARCHAR(4000)
);

DECLARE @fecha DATETIME,
        @nrocomanda VARCHAR(15),
        @moneda CHAR(1),
        @mesa VARCHAR(50),
        @mozo VARCHAR(40);
--PARCHE ICBPER

DECLARE @ICBPER DECIMAL(8, 2);

SELECT TOP 1
       @ICBPER = COALESCE(g.GEN_ICBPER, 0)
FROM dbo.GENERAL g;
--FIN PARCHE

DECLARE @TBLICBPER TABLE
(
    CODART BIGINT,
    ICBPER MONEY
);



IF @xdet IS NULL --IMPRIME TODOS LOS ITEMS DE LA COMANDA
BEGIN
    INSERT INTO @tbldata
    SELECT PEDIDOS.PED_FECHA,
           PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))) AS 'NROCOMANDA',
           PEDIDOS.PED_CANTIDAD,
           PEDIDOS.PED_PRECIO,
           PEDIDOS.PED_IGV,
           PEDIDOS.PED_BRUTO,
           PEDIDOS.PED_HORA,
           PEDIDOS.PED_MONEDA,
           PEDIDOS.PED_SUBTOTAL,
           ARTI.ART_NOMBRE,
           RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON) AS CLI_NOMBRE,
           VEMAEST.VEM_NOMBRE,
           PEDIDOS.PED_OFERTA,
           PEDIDOS.PED_CLIENTE,
           PEDIDOS.PED_FAMILIA2 AS 'PED_FAMILIA',
           ARTI.ART_KEY,
           ARTI.ART_FLAG_STOCK,
           GETDATE(),
           (
               SELECT RTRIM(LTRIM(t.TAB_NOMLARGO))
               FROM dbo.TABLAS t
               WHERE t.TAB_TIPREG = 122
                     AND t.TAB_CODCIA = @CodCia
                     AND t.TAB_NUMTAB = PEDIDOS.PED_FAMILIA2
           ),
           dbo.FnDevuelveCaracteristica(
                                           PEDIDOS.PED_CODCIA,
                                           PEDIDOS.PED_FECHA,
                                           PEDIDOS.PED_NUMFAC,
                                           PEDIDOS.PED_NUMSER,
                                           PEDIDOS.PED_NUMSEC,
                                           PEDIDOS.PED_CODART
                                       )
    FROM dbo.PEDIDOS PEDIDOS
        INNER JOIN dbo.MESAS CLIENTES
            ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
               AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
        INNER JOIN dbo.VEMAEST VEMAEST
            ON PEDIDOS.PED_CODVEN = VEMAEST.VEM_CODVEN
               AND PEDIDOS.PED_CODCIA = VEMAEST.VEM_CODCIA
        INNER JOIN dbo.ARTI ARTI
            ON PEDIDOS.PED_CODART = ARTI.ART_KEY
               AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
    WHERE PEDIDOS.PED_NUMSER = @NumSer
          AND PEDIDOS.PED_NUMFAC = @NumFac
          AND PEDIDOS.PED_CODCIA = @CodCia;

END;
ELSE --SOLO IMPRIME LOS ENVIADOS
BEGIN
    INSERT INTO @tbldata
    SELECT PEDIDOS.PED_FECHA,
           PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))) AS 'NROCOMANDA',
           PEDIDOS.PED_CANTIDAD,
           PEDIDOS.PED_PRECIO,
           PEDIDOS.PED_IGV,
           PEDIDOS.PED_BRUTO,
           PEDIDOS.PED_HORA,
           PEDIDOS.PED_MONEDA,
           PEDIDOS.PED_SUBTOTAL,
           ARTI.ART_NOMBRE,
           RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON) AS CLI_NOMBRE,
           VEMAEST.VEM_NOMBRE,
           PEDIDOS.PED_OFERTA,
           PEDIDOS.PED_CLIENTE,
           PEDIDOS.PED_FAMILIA2 AS 'PED_FAMILIA',
           ARTI.ART_KEY,
           ARTI.ART_FLAG_STOCK,
           GETDATE(),
           (
               SELECT RTRIM(LTRIM(t.TAB_NOMLARGO))
               FROM dbo.TABLAS t
               WHERE t.TAB_TIPREG = 122
                     AND t.TAB_CODCIA = @CodCia
                     AND t.TAB_NUMTAB = PEDIDOS.PED_FAMILIA2
           ),
           dbo.FnDevuelveCaracteristica(
                                           PEDIDOS.PED_CODCIA,
                                           PEDIDOS.PED_FECHA,
                                           PEDIDOS.PED_NUMFAC,
                                           PEDIDOS.PED_NUMSER,
                                           PEDIDOS.PED_NUMSEC,
                                           PEDIDOS.PED_CODART
                                       )
    FROM dbo.PEDIDOS PEDIDOS
        INNER JOIN dbo.MESAS CLIENTES
            ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
               AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
        INNER JOIN dbo.VEMAEST VEMAEST
            ON PEDIDOS.PED_CODVEN = VEMAEST.VEM_CODVEN
               AND PEDIDOS.PED_CODCIA = VEMAEST.VEM_CODCIA
        INNER JOIN dbo.ARTI ARTI
            ON PEDIDOS.PED_CODART = ARTI.ART_KEY
               AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
    WHERE PEDIDOS.PED_NUMSER = @NumSer
          AND PEDIDOS.PED_NUMFAC = @NumFac
          AND PEDIDOS.PED_CODCIA = @CodCia
          AND PEDIDOS.PED_CODART IN
              (
                  SELECT parametro FROM dbo.FnTextoaTabla(@xdet)
              )
          AND PEDIDOS.PED_NUMSEC IN
              (
                  SELECT parametro FROM dbo.FnTextoaTabla(@xnumsec)
              )
    ORDER BY PEDIDOS.PED_FECHAREG;

    --obtengo datos para impresion
    SELECT @fecha = PEDIDOS.PED_FECHA,
           @nrocomanda = PEDIDOS.PED_NUMSER + '-' + RTRIM(LTRIM(STR(PEDIDOS.PED_NUMFAC))),
           @moneda = PEDIDOS.PED_MONEDA,
           @mesa = RTRIM(LTRIM(CLIENTES.MES_DESCRIP)) + ' - ' + dbo.FnDevuelveZona(@CodCia, CLIENTES.MES_CODZON),
           @mozo = VEMAEST.VEM_NOMBRE
    FROM dbo.PEDIDOS PEDIDOS
        INNER JOIN dbo.MESAS CLIENTES
            ON PEDIDOS.PED_CODCLIE = CLIENTES.MES_CODMES
               AND PEDIDOS.PED_CODCIA = CLIENTES.MES_CODCIA
        INNER JOIN dbo.VEMAEST VEMAEST
            ON PEDIDOS.PED_CODVEN = VEMAEST.VEM_CODVEN
               AND PEDIDOS.PED_CODCIA = VEMAEST.VEM_CODCIA
        INNER JOIN dbo.ARTI ARTI
            ON PEDIDOS.PED_CODART = ARTI.ART_KEY
               AND PEDIDOS.PED_CODCIA = ARTI.ART_CODCIA
    WHERE PEDIDOS.PED_NUMSER = @NumSer
          AND PEDIDOS.PED_NUMFAC = @NumFac
          AND PEDIDOS.PED_CODCIA = @CodCia
          AND PEDIDOS.PED_CODART IN
              (
                  SELECT parametro FROM dbo.FnTextoaTabla(@xdet)
              )
          AND PEDIDOS.PED_NUMSEC IN
              (
                  SELECT parametro FROM dbo.FnTextoaTabla(@xnumsec)
              );


END;

--actualizo los pedidos de acuerdo a impresion
UPDATE PEDIDOS
SET PED_APROBADO = '1'
WHERE PEDIDOS.PED_CODART IN
      (
          SELECT parametro FROM dbo.FnTextoaTabla(@xdet)
      )
      AND PEDIDOS.PED_NUMSEC IN
          (
              SELECT parametro FROM dbo.FnTextoaTabla(@xnumsec)
          );


--tabla para los productos que son combos
DECLARE @tblcombos TABLE
(
    codcombo BIGINT,
    cant BIGINT,
    sec TINYINT IDENTITY(1, 1),
    hora VARCHAR(15)
);

INSERT INTO @tblcombos
SELECT codprod,
       PED_CANTIDAD,
       PED_HORA
FROM @tbldata
WHERE flag = 'C';


--SELECT * FROM @tblcombos
DECLARE @hora VARCHAR(15);

IF EXISTS (SELECT TOP 1 codcombo FROM @tblcombos)
BEGIN --ENTRA AQUI ES PORQUE TIENE COMBOS

    DECLARE @tbltmpCombos TABLE
    (
        PED_FECHA DATETIME,
        NROCOMANDA VARCHAR(10),
        PED_CANTIDAD MONEY,
        PED_PRECIO MONEY,
        PED_IGV MONEY,
        PED_BRUTO MONEY,
        PED_HORA VARCHAR(15),
        PED_MONEDA CHAR(1),
        PED_SUBTOTAL MONEY,
        ART_NOMBRE VARCHAR(80),
        CLI_NOMBRE VARCHAR(80),
        VEM_NOMBRE VARCHAR(60),
        ped_oferta VARCHAR(300),
        ped_cliente VARCHAR(120),
        ped_familia2 INT,
        codprod BIGINT,
        flag CHAR(1),
        FAMILIA VARCHAR(100),
        num TINYINT IDENTITY
    );

    INSERT INTO @tbltmpCombos
    SELECT PED_FECHA,
           NROCOMANDA,
           PED_CANTIDAD,
           PED_PRECIO,
           PED_IGV,
           PED_BRUTO,
           PED_HORA,
           PED_MONEDA,
           PED_SUBTOTAL,
           ART_NOMBRE,
           CLI_NOMBRE,
           VEM_NOMBRE,
           PED_OFERTA,
           PED_CLIENTE,
           ped_familia,
           codprod,
           flag,
           FAMILIA
    FROM @tbldata
    WHERE flag = 'C';


    DELETE FROM @tbldata
    WHERE flag = 'C';

    DECLARE @codcombo BIGINT,
            @cant BIGINT,
            @num TINYINT,
            @po VARCHAR(300);
    --aqui entra el cursor
    DECLARE cCombos CURSOR FOR
    SELECT codcombo,
           cant,
           sec,
           hora
    FROM @tblcombos;

    OPEN cCombos;

    FETCH cCombos
    INTO @codcombo,
         @cant,
         @num,
         @hora;

    WHILE (@@Fetch_Status = 0)
    BEGIN

        INSERT INTO @tbldata
        SELECT PED_FECHA,
               NROCOMANDA,
               PED_CANTIDAD,
               PED_PRECIO,
               PED_IGV,
               PED_BRUTO,
               PED_HORA,
               PED_MONEDA,
               PED_SUBTOTAL,
               ART_NOMBRE,
               CLI_NOMBRE,
               VEM_NOMBRE,
               ped_oferta,
               ped_cliente,
               ped_familia2,
               codprod,
               flag,
               GETDATE(),
               FAMILIA,
               ''
        FROM @tbltmpCombos
        WHERE codprod = @codcombo
              AND num = @num;


        INSERT INTO @tbldata
        SELECT @fecha,
               @nrocomanda,
               pa.PA_PROM * @cant, --ped_cantidad  GTS ACA CANTIDAD DE COMBOS
               0,
               0,
               0,
               @hora,
               @moneda,
               0,
               ' * ' + ar.ART_NOMBRE,
               @mesa,
               @mozo,
               ISNULL(@po, ''),
               '',
               ar.ART_FAMILIA,
               pa.PA_CODART,
               ar.ART_FLAG_STOCK,
               GETDATE(),
               (
                   SELECT RTRIM(LTRIM(t.TAB_NOMLARGO))
                   FROM dbo.TABLAS t
                   WHERE t.TAB_TIPREG = 122
                         AND t.TAB_CODCIA = @CodCia
                         AND t.TAB_NUMTAB = ar.ART_FAMILIA
               ),
               ''
        FROM PAQUETES pa
            INNER JOIN ARTI ar
                ON pa.PA_CODCIA = ar.ART_CODCIA
                   AND PA_CODART = ar.ART_KEY
        WHERE pa.PA_CODPA = @codcombo
              AND pa.PA_CODCIA = @CodCia;

        FETCH cCombos
        INTO @codcombo,
             @cant,
             @num,
             @hora;
    END;

    CLOSE cCombos;
    DEALLOCATE cCombos;


END;


--PARCHE PARA AGREGAR FAMILIA AL COMBO FALTANTE
--DECLARE @MIN INT ,
--    @MAX INT
--DECLARE @TBLFAMILIA TABLE
--    (
--      IDFAMILIA INT ,
--      INDICE INT IDENTITY
--    )
--INSERT  INTO @TBLFAMILIA
--        ( IDFAMILIA
--        )
--        SELECT DISTINCT
--                PED_FAMILIA
--        FROM    @TBLDATA

--SELECT  @MIN = MIN(T.INDICE)
--FROM    @TBLFAMILIA t
--SELECT  @MAX = MAX(T.INDICE)
--FROM    @TBLFAMILIA t

--WHILE @MIN <= @MAX
--    BEGIN
--        IF NOT EXISTS ( SELECT TOP 1
--                                NROCOMANDA
--                        FROM    @TBLDATA
--                        WHERE   PED_FAMILIA = ( SELECT TOP 1
--                                                  T.IDFAMILIA
--                                                FROM
--                                                  @TBLFAMILIA t
--                                                WHERE
--                                                  T.INDICE = @MIN
--                                              )
--                                AND flag = 'C' )
--            BEGIN
--            --SELECT 'noexiste'

--                INSERT  INTO @TBLDATA
--                        SELECT  PED_FECHA ,
--                                NROCOMANDA ,
--                                PED_CANTIDAD ,
--                                PED_PRECIO ,
--                                PED_IGV ,
--                          PED_BRUTO ,
--                                dbo.FnDevuelveHora(GETDATE()) ,
--                                PED_MONEDA ,
--                                PED_SUBTOTAL ,
--                                ART_NOMBRE ,
--                                CLI_NOMBRE ,
--                                VEM_NOMBRE ,
--                                PED_OFERTA ,
--                                ped_cliente ,
--                                ( SELECT TOP 1
--                                            T.IDFAMILIA
--                                  FROM      @TBLFAMILIA t
--                                  WHERE     T.INDICE = @MIN
--                                ) ,
--                                codprod ,
--                                'C' ,
--                                actual ,
--                                ( SELECT TOP 1
--                                            t.tab_nomlargo
--                                  FROM      dbo.TABLAS t
--                                  WHERE     t.TAB_TIPREG = 122
--                                            AND t.TAB_NUMTAB = ( SELECT TOP 1
--                                                  T.IDFAMILIA
--                                                  FROM
--                                                  @TBLFAMILIA t
--                                                  WHERE
--                                                  T.INDICE = @MIN
--                                                  )
--                                ) ,
--                                ''
--                        FROM    @TBLDATA
--                        WHERE   PED_FAMILIA = ( SELECT TOP 1
--                                                  t.ped_familia
--                                                FROM
--                                                  @tbldata t
--                                                WHERE
--                                                  T.flag = 'C'
--                                              )
--                                AND flag = 'C'


--            END


--        SET @MIN = @MIN + 1
--    END

IF @precuenta = 1 --precuenta
BEGIN
    SELECT *
    FROM @tbldata
    ORDER BY flag,
             PED_HORA;
END;
ELSE --comanda
BEGIN
    SELECT *
    FROM @tbldata
    ORDER BY ped_familia,
             flag,
             PED_HORA;
END;






GO