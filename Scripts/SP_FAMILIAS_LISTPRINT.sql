/*
exec SP_FAMILIAS_LISTPRINT '01'
*/
ALTER PROC SP_FAMILIAS_LISTPRINT @CODCIA CHAR(2)
AS
    SET NOCOUNT ON 
    
    DECLARE @TBLGRUPOS TABLE
        (
          INDICE INT IDENTITY ,
          ITEM CHAR(1)
        )
    DECLARE @TBLFINAL TABLE
        (
          FAMILIA VARCHAR(100) ,
          IMPRESORA VARCHAR(100),
          IMPRESORA2 VARCHAR(100)
        )
    DECLARE @TBLTEMPORAL TABLE
        (
          INDICE INT IDENTITY ,
          IDFAMILIA INT
        )
            
    INSERT  INTO @TBLGRUPOS
            ( ITEM
            )
            SELECT DISTINCT
                    T.GRUPO
            FROM    dbo.TABLAS t
            WHERE   T.TAB_CODCIA = @CODCIA
                    AND T.TAB_TIPREG = 122
                    AND GRUPO IS NOT NULL
                    AND IMPRESORA IS NOT NULL
            
    DECLARE @FAMILIAS VARCHAR(100) ,
        @IDFAMILIA INT
    DECLARE @MIN INT ,
        @MIN1 INT ,
        @MAX1 INT ,
        @MAX INT ,
        @GRUPO CHAR(1) ,
        @IMPRESORA VARCHAR(50),@IMPRESORA2 VARCHAR(50)
    SELECT  @MIN = MIN(T.INDICE)
    FROM    @TBLGRUPOS t
    SELECT  @MAX = MAX(T.INDICE)
    FROM    @TBLGRUPOS t
            
    WHILE @MIN <= @MAX
        BEGIN
            SELECT  @GRUPO = T.ITEM
            FROM    @TBLGRUPOS t
            WHERE   T.INDICE = @MIN
            
          -- SET @FAMILIAS = @FAMILIAS + (
            INSERT  INTO @TBLTEMPORAL
                    ( IDFAMILIA
                    )
                    SELECT  CAST(T.TAB_NUMTAB AS VARCHAR(20))
                    FROM    dbo.TABLAS t
                    WHERE   T.TAB_CODCIA = @CODCIA
                            AND T.IMPRESORA IS NOT NULL
                            AND T.TAB_TIPREG = 122
                            AND T.GRUPO = @GRUPO
                    --)
                    
                    
            SELECT  @MIN1 = MIN(T.INDICE)
            FROM    @TBLTEMPORAL t
            SELECT  @MAX1 = MAX(T.INDICE)
            FROM    @TBLTEMPORAL t
    
            SET @FAMILIAS = ''
            WHILE @MIN1 <= @MAX1
                BEGIN
                    SELECT  @IDFAMILIA = T.IDFAMILIA
                    FROM    @TBLTEMPORAL t
                    WHERE   T.INDICE = @MIN1
                    SELECT TOP 1
                            @IMPRESORA = T.IMPRESORA,@IMPRESORA2 = COALESCE(t.IMPRESORA2,'')
                    FROM    dbo.TABLAS t
                    WHERE   T.TAB_CODCIA = @CODCIA
                            AND T.TAB_TIPREG = 122
                            AND TAB_NUMTAB = @IDFAMILIA
                    SET @FAMILIAS = @FAMILIAS
                        + CAST(@IDFAMILIA AS VARCHAR(10)) + '|'
					
                    SET @MIN1 = @MIN1 + 1
                END
            INSERT  INTO @TBLFINAL
                    ( FAMILIA, IMPRESORA ,IMPRESORA2)
                    SELECT  @FAMILIAS ,
                            @IMPRESORA,@IMPRESORA2
  
            DELETE  FROM @TBLTEMPORAL
            SET @MIN = @MIN + 1
        END
            
    SELECT  *
    FROM    @TBLFINAL t
GO