ALTER PROC [dbo].[SP_FAMILIA_REGISTRAR]
    (
      @CODCIA CHAR(2) ,
      @DENOMINACION VARCHAR(50) ,
      @IMPRESORA VARCHAR(50) ,
      @IMPRESORA2 VARCHAR(50),
      @GRUPO CHAR(1),
      @DSCTO MONEY,
      @VISIBLE BIT
    )
AS
    SET NOCOUNT ON 

    DECLARE @CODIGO INT


    IF NOT EXISTS ( SELECT TOP 1
                            T.TAB_TIPREG
                    FROM    dbo.TABLAS t
                    WHERE   T.TAB_TIPREG = 122
                            AND T.TAB_CODCIA = @CODCIA
                            AND T.TAB_NOMLARGO = @DENOMINACION )
        BEGIN

            SELECT  @CODIGO = ISNULL(MAX(T.TAB_NUMTAB), 0) + 1
            FROM    dbo.TABLAS t
            WHERE   T.TAB_TIPREG = 122
                    AND T.TAB_CODCIA = @CODCIA

            INSERT  INTO dbo.TABLAS
                    ( TAB_CODCIA ,
                      TAB_TIPREG ,
                      TAB_NUMTAB ,
                      TAB_NOMLARGO ,
                      TAB_NOMCORTO ,
                      TAB_CODART ,
                      TAB_CONTABLE2 ,
                      TAB_FECHA_CONTROL ,
                      IMPRESORA ,
                      GRUPO,tab_descuento,TAB_VISIBLE,IMPRESORA2
	                )
            VALUES  ( @CODCIA , -- TAB_CODCIA - char(2)
                      122 , -- TAB_TIPREG - int
                      @CODIGO , -- TAB_NUMTAB - int
                      @DENOMINACION , -- TAB_NOMLARGO - char(40)
                      LEFT(@DENOMINACION, 10) , -- TAB_NOMCORTO - char(10)
                      0 , -- TAB_CODART - numeric
                      '0' , -- TAB_CONTABLE2 - varchar(50)
                      GETDATE() , -- TAB_FECHA_CONTROL - datetime
                      @IMPRESORA , -- IMPRESORA - varchar(50)
                      @GRUPO  -- GRUPO - char(1)
                      ,@DSCTO,@VISIBLE,@IMPRESORA2
	                )
        END
    ELSE
        BEGIN
            RAISERROR('La Descripción Proporcionada ya existe',16,1)

        END
GO