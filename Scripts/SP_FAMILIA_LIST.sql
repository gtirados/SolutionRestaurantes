IF EXISTS ( SELECT  *
            FROM    INFORMATION_SCHEMA.routines
            WHERE   routine_type = 'PROCEDURE'
                    AND ROUTINE_NAME = 'SP_FAMILIA_LIST' )
    DROP PROCEDURE SP_FAMILIA_LIST
GO
/*
SP_FAMILIA_MODIFICAR '01'
*/
CREATE PROC SP_FAMILIA_LIST ( @CODCIA CHAR(2) )
AS
    SET NOCOUNT ON 

    SELECT  '-1' AS 'COD' ,
            '.: SELECCIONE :.' AS 'DEN'
    UNION
    SELECT  T.TAB_NUMTAB AS 'COD' ,
            LTRIM(RTRIM(T.TAB_NOMLARGO)) AS 'DEN'
    FROM    dbo.TABLAS t
    WHERE   T.TAB_TIPREG = 122

GO
