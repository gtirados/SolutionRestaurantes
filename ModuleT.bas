Attribute VB_Name = "ModuleT"
Dim cnDBF As ADODB.Connection
Dim rs As ADODB.Recordset

'PROCESO DE TRANSF. DE DATOS
Public Sub transferenciaDatos()
Dim sConexion As String
Dim sSql As String
Dim sName As String
Dim sTabla As String
Dim i As Integer
Dim NumeroDoc As String
Dim CodClie As String
Dim codtmp As String
Dim RSVend As New ADODB.Recordset
Dim sCodClie As Long
Dim Importe As Currency

    Set cnDBF = New ADODB.Connection
    cnDBF.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DSN_DBF"
    
        Set rs = New ADODB.Recordset
        sSql = "select * from cargav "
        rs.Open sSql, cnDBF
        
        If Not rs.EOF Then
            Do While Not rs.EOF
                If IsNull(rs("COD_CLIe")) Then GoTo OTRO
                SQ_OPER = 4
                PUB_RUC = rs("COD_CLIe")
                pu_codcia = LK_CODCIA
                pu_cp = PUB_CP
                LEER_CLI_LLAVE
                If Not cli_ruc.EOF Then
                    sCodClie = cli_ruc("cli_codclie")
                Else
                    MsgBox "C/P no existe " & rs("COD_CLIe")
                    Debug.Print rs("nro_fact")
                    GoTo OTRO
                End If
                FORMGEN.i_codcli.Text = sCodClie
                FORMGEN.i_codcli_KeyPress 13
                FORMGEN.i_codcli_LostFocus
                
                
                fectmp = Replace(rs("FECHa_f"), "Jan", "Ene")
                fectmp = Replace(fectmp, "Dec", "Dic")
                FEC = Format(fectmp, "dd/mm/yy")
                'FEC = "15/01/04"
                FORMGEN.i_fecha_compra.Text = FEC
                FORMGEN.i_ds.Text = IIf(rs("moneda") = 2, "D", "S")
                If rs("moneda") = 2 Then
                    If Val(rs("cambio")) = 305 Then
                        Importe = (Val(rs("total")) - Val(rs("monto_c"))) / Val(rs("cambio")) / 3.5
                    Else
                        Importe = (Val(rs("total")) - Val(rs("monto_c"))) / Val(rs("cambio"))
                    End If
                Else
                    Importe = rs("total") - rs("monto_c")
                End If
                
                Importe = Format(Importe, "0.00")
                FORMGEN.i_importe_amort.Text = Importe 'rs("total")
                FORMGEN.i_fbg.Text = IIf(rs("fact_bol") = "Verdadero", "F", "B") 'falta definir
                FORMGEN.i_numser_c.Text = Left(rs("nro_fact"), 3)
                NumeroDoc = rs("Nro_FACt")
                pos = InStr(1, NumeroDoc, "-")
                NumeroDoc = Mid(NumeroDoc, pos + 1, Len(NumeroDoc) - pos)
                
                FORMGEN.i_numfac_c.Text = NumeroDoc
                FORMGEN.i_fecha_vcto.Text = Format(DateAdd("d", rs("dias_letra"), FEC), "dd/mm/yyyy") 'falta definir si es correcto
                FORMGEN.i_concepto.Text = rs("cliente")
                FORMGEN.i_fecha_can.Text = Format(FEC, "dd/mm/yy")
                
                codven = IIf(IsNull(rs("cod_vend")), "", rs("cod_vend"))
                
                RSVend.Open "Select * from Vendedor where CODIGO = '" & codven & "'", cnDBF
                If Not RSVend.EOF Then
                    codven = IIf(IsNull(RSVend("key_v")), 0, RSVend("key_v"))
                Else
                    codven = 0
                End If
                RSVend.Close
        
                FORMGEN.i_codven = codven
                Call FORMGEN.grabar_Click
OTRO:
                rs.MoveNext
            Loop
        End If
        rs.Close
End Sub
Public Sub transferenciaDatos1()
Dim sConexion As String
Dim sSql As String
Dim sName As String
Dim sTabla As String
Dim i As Integer
Dim NumeroDoc As String
Dim CodClie As String
Dim codtmp As String
Dim RSVend As New ADODB.Recordset
Dim sCodClie As Long
Dim Importe As Currency

    Set cnDBF = New ADODB.Connection
    cnDBF.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DSN_DBF"
    
        Set rs = New ADODB.Recordset
        sSql = "select * from cargac "
        rs.Open sSql, cnDBF
        
        If Not rs.EOF Then
            Do While Not rs.EOF
                If IsNull(rs("codigo")) Then GoTo OTRO
                SQ_OPER = 4
                PUB_RUC = rs("codigo")
                pu_codcia = LK_CODCIA
                pu_cp = PUB_CP
                LEER_CLI_LLAVE
                If Not cli_ruc.EOF Then
                    sCodClie = cli_ruc("cli_codclie")
                Else
                    MsgBox "C/P no existe " & rs("COD_CLIe")
                    Debug.Print rs("nro_fact")
                    GoTo OTRO
                End If
                FORMGEN.i_codcli.Text = sCodClie
                FORMGEN.i_codcli_KeyPress 13
                FORMGEN.i_codcli_LostFocus
                
                fectmp = Replace(rs("FECHa_f"), "Jan", "Ene")
                fectmp = Replace(fectmp, "Dec", "Dic")
                FEC = Format(fectmp, "dd/mm/yy")
                
                'FEC = "15/01/04"
                FORMGEN.i_fecha_compra.Text = FEC
                FORMGEN.i_ds.Text = IIf(rs("moneda") = 2, "D", "S")
                Importe = Val(rs("total")) - Val(rs("monto_c"))
                
                
                Importe = Format(Importe, "0.00")
                FORMGEN.i_importe_amort.Text = Importe 'rs("total")
                FORMGEN.i_fbg.Text = "F" 'Left(rs("fb"), 1) 'falta definir
                FORMGEN.i_numser_c.Text = Left(rs("nro_fact"), 3)
                NumeroDoc = rs("Nro_FACt")
                pos = InStr(1, NumeroDoc, "-")
                NumeroDoc = Mid(NumeroDoc, pos + 1, Len(NumeroDoc) - pos)
                
                fecV = Replace(rs("fecha_v"), "Jan", "Ene")
                fecV = Replace(rs("fecha_v"), "Dec", "Dic")
                fecV = Replace(rs("fecha_v"), "Apr", "Abr")
                If fecV = "  -   -" Then
                    fecV = FEC
                End If
                fecV = Format(fecV, "dd/mm/yy")
                
                FORMGEN.i_numfac_c.Text = NumeroDoc
                FORMGEN.i_fecha_vcto.Text = Format(fecV, "dd/mm/yyyy") 'falta definir si es correcto
                FORMGEN.i_concepto.Text = rs("proveedor")
                FORMGEN.i_fecha_can.Text = Format(FEC, "dd/mm/yy")
                FORMGEN.i_codven = 6
                Call FORMGEN.grabar_Click
OTRO:
                rs.MoveNext
            Loop
        End If
        rs.Close
End Sub
Public Sub Levanta2401()
Dim Sec_Serie As Integer
Dim Sec_Numero As Integer
Dim Cnn_DBF As ADODB.Connection
    
    pub_mensaje = "Generar los Documentos ¿Desea Continuar... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
        Exit Sub
    End If
    sCnnDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DSN_DBF" '"Provider=SQLOLEDB.1;Persist Security Info=False;pwd=;User ID=sa;Initial Catalog=BDHUEM;Data Source=PC01" '"Provider=MSDASQL;Data Source=DSN_TMP"
    Cnn_DBF.CursorLocation = adUseClient
    Cnn_DBF.Open sCnnDBF
    pub_cadena = "SELECT  * FROM carga2401 WHERE CODCIA = '" & LK_CODCIA & "' and fecha='" & LK_FECHA_DIA & "'"
    rs.Open s_Sql, Cnn_DBF, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    With FORMGEN
    Do While Not rs.EOF
        DoEvents
        PUB_PROCESO = 0
        'llena para grabar 2401
        If numfactmp <> PUB_NUMFAC And numfactmp <> 0 Then
            .grabar_Click
        End If
        If numfactmp <> PUB_NUMFAC Then
            fila = 2
            For fila = 0 To .i_def.ListCount - 1
              .i_def.ListIndex = fila
              If Val(Left(i_def.Text, 2)) = Val(rs!ped_CONDI) Then
                 Exit For
              End If
            Next fila
            
            .i_def_LostFocus
            .i_codcli.Text = rs!PED_CODCLIE
            .i_codcli_KeyPress 13
            .i_codcli_LostFocus
            .i_codven.Text = rs!PED_CODVEN
            .i_codven_KeyPress 13
            .i_ds.Text = rs("Moneda")
            If rs!ped_FBG = "F" Then
              .i_fbg.ListIndex = 0
            Else
              .i_fbg.ListIndex = 1
            End If
            .i_codcho.Text = rs("chofer")
            .i_fecha_compra.Text = rs("fecemi")
            .i_dias.Text = rs!ped_DIAS
            .i_cambio.Value = 1
            .i_cambio.Visible = True
            .i_numser.Text = PUB_NUMSER
            .i_numfac.Text = PUB_NUMFAC
            .i_TEXTONCRE.Text = PUB_NUMFAC
            .i_dircli_GotFocus
            .i_gastos.Text = rs("desctgrl")
            PUB_PEDSER = rs!PED_NUMSER
            PUB_PEDFAC = rs!PED_NUMFAC
        End If
            
         grid_fac.Rows = fila + 2
         SQ_OPER = 3
         pu_alterno = rs!PED_CODART
         pu_codcia = LK_CODCIA
         LEER_ART_LLAVE
         If art_llave_alt.EOF Then
              MsgBox "Articulo no existe: " & art_llave_alt("art_alterno")
              GoTo OTRO
         End If
         PUB_KEY = art_llave_alt("art_codart")
         pu_codcia = LK_CODCIA
         SQ_OPER = 1
         LEER_ART_LLAVE
         SQ_OPER = 1
         PUB_CODART = PUB_KEY
         pu_codcia = LK_CODCIA
         LEER_ARM_LLAVE
         If arm_llave.EOF Then
            MsgBox "Datos de Articulos ...errados..Revisar"
            Exit Sub
         End If
         SQ_OPER = 3
         PUB_UNIDADS = rs("Unidad")
         LEER_PRE_LLAVE
         grid_fac.TextMatrix(fila, 34) = redondea(pre_llave!pre_PESO * rs!PED_CANTIDAD)
         grid_fac.TextMatrix(fila, 37) = pre_llave!pre_PESO
         grid_fac.TextMatrix(fila, 43) = Nulo_Valor0(pre_llave!PRE_LITRO)
         grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
         grid_fac.TextMatrix(fila, 1) = PUB_CODART
         grid_fac.TextMatrix(fila, 16) = PUB_CODART
         grid_fac.TextMatrix(fila, 2) = 0
         grid_fac.TextMatrix(fila, 3) = rs!PED_UNIDAD
         grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
         If Nulo_Valor0(rs!PED_EQUIV) > 0 Then
         grid_fac.TextMatrix(fila, 4) = rs!PED_CANTIDAD / Nulo_Valor0(rs!PED_EQUIV)
         Else
         grid_fac.TextMatrix(fila, 4) = rs!PED_CANTIDAD
         End If
         grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(rs!PED_EQUIV)
         grid_fac.TextMatrix(fila, 5) = Nulo_Valors(rs!PED_UNIDAD)
         grid_fac.TextMatrix(fila, 6) = Nulo_Valors(rs!PED_PRECIO)
         grid_fac.TextMatrix(fila, 8) = 0
         grid_fac.TextMatrix(fila, 10) = rs!ped_DESCTO_pre
         grid_fac.TextMatrix(fila, 42) = rs!ped_DESCTO
         grid_fac.TextMatrix(fila, 12) = SUT_LLAVE!SUT_SIGNO_ARM
         grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
         grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
         grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
         grid_fac.TextMatrix(fila, 26) = 10 'rs!ped_TIPMOV
         grid_fac.TextMatrix(fila, 27) = LK_CODCIA
         grid_fac.TextMatrix(fila, 28) = PUB_NUMSER
         grid_fac.TextMatrix(fila, 29) = rs!ped_FBG
         grid_fac.TextMatrix(fila, 30) = PUB_NUMFAC
         grid_fac.TextMatrix(fila, 31) = fila - 1
         grid_fac.TextMatrix(fila, 33) = PUB_CODART
         fila = fila + 1
        numfactmp = PUB_NUMFAC
OTRO:
        rs.MoveNext
    Loop
    End With
    rs.Close
    Cnn_DBF.Close
End Sub

Public Sub Levanta1401()

Dim Sec_Serie As Integer
Dim Sec_Numero As Integer
Dim Cnn_DBF As ADODB.Connection
    
    pub_mensaje = "Generar los Documentos ¿Desea Continuar... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
        Exit Sub
    End If
    sCnnDBF = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DSN_DBF" '"Provider=SQLOLEDB.1;Persist Security Info=False;pwd=;User ID=sa;Initial Catalog=BDHUEM;Data Source=PC01" '"Provider=MSDASQL;Data Source=DSN_TMP"
    Cnn_DBF.CursorLocation = adUseClient
    Cnn_DBF.Open sCnnDBF
    pub_cadena = "SELECT  * FROM carga1401 WHERE CODCIA = '" & LK_CODCIA & "' and fecha='" & LK_FECHA_DIA & "'"
    rs.Open s_Sql, Cnn_DBF, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    With FORMGEN
    Do While Not rs.EOF
        DoEvents
        PUB_PROCESO = 0
        'llena para grabar 1401
        If numfactmp <> NumDoc And numserTmp <> serdoc And numfactmp <> 0 And serdoc <> 0 Then
            .grabar_Click
        End If
        If numfactmp <> NumDoc Then
            fila = 2
            For fila = 0 To .i_def.ListCount - 1
              .i_def.ListIndex = fila
              If Val(Left(i_def.Text, 2)) = Val(rs!ped_CONDI) Then
                 Exit For
              End If
            Next fila
            .i_def_LostFocus
            .i_codcli.Text = rs!PED_CODCLIE
            .i_codcli_KeyPress 13
            .i_codcli_LostFocus
            .i_ds.Text = rs("Moneda")
            .i_fecha_compra = rs("fechaemision")
            .i_numser_c = rs("serfac")
            .i_numfac_c = rs("numfac")
            .i_serguia = rs("serguia")
            .i_numguia = rs("numguia")
            .i_dias = rs("dias")
            If rs!ped_FBG = "F" Then
              .i_fbg.ListIndex = 0
            Else
              .i_fbg.ListIndex = 1
            End If
            If rs("numfac") <> 0 And rs("numser") <> 0 Then
                NumDoc = rs("numfac")
                numser = rs("serfac")
            Else
                NumDoc = rs("numguia")
                numser = rs("serguia")
            End If
            If NumDoc = 0 Or numser = 0 Then
                MsgBox ("error con los numero de documentos")
            End If
        End If
         grid_fac.Rows = fila + 2
         SQ_OPER = 3
         pu_alterno = rs!PED_CODART
         pu_codcia = LK_CODCIA
         LEER_ART_LLAVE
         If art_llave_alt.EOF Then
              MsgBox "Articulo no existe: " & art_llave_alt("art_alterno")
              GoTo OTRO
         End If
         PUB_KEY = art_llave_alt("art_codart")
         pu_codcia = LK_CODCIA
         SQ_OPER = 1
         LEER_ART_LLAVE
         SQ_OPER = 1
         PUB_CODART = PUB_KEY
         pu_codcia = LK_CODCIA
         LEER_ARM_LLAVE
         If arm_llave.EOF Then
            MsgBox "Datos de Articulos ...errados..Revisar"
            Exit Sub
         End If
         SQ_OPER = 3
         PUB_UNIDADS = rs("Unidad")
         LEER_PRE_LLAVE
         grid_fac.TextMatrix(fila, 34) = redondea(pre_llave!pre_PESO * rs!PED_CANTIDAD)
         grid_fac.TextMatrix(fila, 37) = pre_llave!pre_PESO
         grid_fac.TextMatrix(fila, 43) = Nulo_Valor0(pre_llave!PRE_LITRO)
         grid_fac.TextMatrix(fila, 0) = art_LLAVE!art_nombre
         grid_fac.TextMatrix(fila, 1) = PUB_CODART
         grid_fac.TextMatrix(fila, 16) = PUB_CODART
         grid_fac.TextMatrix(fila, 2) = 0
         grid_fac.TextMatrix(fila, 3) = rs!PED_UNIDAD
         grid_fac.TextMatrix(fila, 11) = arm_llave!arm_cospro
         If Nulo_Valor0(rs!PED_EQUIV) > 0 Then
         grid_fac.TextMatrix(fila, 4) = rs!PED_CANTIDAD / Nulo_Valor0(rs!PED_EQUIV)
         Else
         grid_fac.TextMatrix(fila, 4) = rs!PED_CANTIDAD
         End If
         grid_fac.TextMatrix(fila, 14) = Nulo_Valor0(rs!PED_EQUIV)
         grid_fac.TextMatrix(fila, 5) = Nulo_Valors(rs!PED_UNIDAD)
         grid_fac.TextMatrix(fila, 6) = Nulo_Valors(rs!PED_PRECIO)
         grid_fac.TextMatrix(fila, 8) = 0
         grid_fac.TextMatrix(fila, 10) = rs!ped_DESCTO_pre
         grid_fac.TextMatrix(fila, 42) = rs!ped_DESCTO
         grid_fac.TextMatrix(fila, 12) = SUT_LLAVE!SUT_SIGNO_ARM
         grid_fac.TextMatrix(fila, 21) = Nulo_Valors(art_LLAVE!art_flag_stock)
         grid_fac.TextMatrix(fila, 23) = Nulo_Valors(art_LLAVE!ART_EX_IGV)
         grid_fac.TextMatrix(fila, 24) = Nulo_Valor0(art_LLAVE!ART_POR_IGV)
         grid_fac.TextMatrix(fila, 26) = 20 'rs!ped_TIPMOV
         grid_fac.TextMatrix(fila, 27) = LK_CODCIA
         grid_fac.TextMatrix(fila, 28) = PUB_NUMSER
         grid_fac.TextMatrix(fila, 29) = rs!ped_FBG
         grid_fac.TextMatrix(fila, 30) = PUB_NUMFAC
         grid_fac.TextMatrix(fila, 31) = fila - 1
         grid_fac.TextMatrix(fila, 33) = PUB_CODART
         fila = fila + 1
        numfactmp = NumDoc
        numserTmp = serdoc
OTRO:
        rs.MoveNext
    Loop
    End With
    rs.Close
    Cnn_DBF.Close
End Sub
