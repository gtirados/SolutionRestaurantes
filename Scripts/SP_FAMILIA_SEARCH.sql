/*
SP_FAMILIA_SEARCH '01'
*/
ALTER PROC SP_FAMILIA_SEARCH
    (
      @CODCIA CHAR(2) ,
      @SEARCH VARCHAR(50) = NULL
    )
AS
    SET NOCOUNT ON 
    
    SELECT  T.TAB_NUMTAB AS 'IDE' ,
            T.TAB_NOMLARGO AS 'NOM' ,
            COALESCE(T.IMPRESORA, '') AS 'PRINT' ,
            COALESCE(T.IMPRESORA2,'') AS 'PRINT2',
            COALESCE(T.GRUPO, '') AS 'GRUPO' ,
            COALESCE(T.TAB_DESCUENTO, 0) AS 'DSCTO' ,
            CASE WHEN ISNULL(T.TAB_VISIBLE, 0) = 1 THEN 'SI'
                 ELSE 'NO'
            END AS 'VISIBLE'
    FROM    dbo.TABLAS t
    WHERE   t.TAB_TIPREG = 122
            AND T.TAB_CODCIA = @CODCIA
            AND ( T.TAB_NOMLARGO LIKE '%' + @SEARCH + '%'
                  OR @SEARCH IS NULL
                )
GO