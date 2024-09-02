Attribute VB_Name = "modImpresion"

Public Sub ImprimirDocumentoVenta(xCodTipoDocto As String, _
                                  xTipoDocto, _
                                  xEsconsumo As Boolean, _
                                  xSerie As String, _
                                  xNumero As Double, _
                                  xTotal As Double, _
                                  xSubTotal As Double, _
                                  xIgv As Double, _
                                  xDireccion As String, _
                                  xRuc As String, _
                                  xcliente As String, _
                                  xdni As String, _
                                  xCia As String, _
                                  xICBPER As Double, _
                                  xEsprom As Boolean, _
                                  xNroCopias As Integer)

    'RECUPERANDO EL NOMBRE DEL ARCHIVO
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ARCHIVO_PRINT"
    oCmdEjec.CommandType = adCmdStoredProc
    
    Dim ORSd        As ADODB.Recordset

    Dim RutaReporte As String

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adChar, adParamInput, 2, xCodTipoDocto)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COMSUMO", adBoolean, adParamInput, , xEsconsumo)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, xCia)
    
    Set ORSd = oCmdEjec.Execute
    
    RutaReporte = PUB_RUTA_REPORTE & ORSd!ReportE

    'OBTENIENDO DATOS DEL CLIENTE
    '    LimpiaParametros oCmdEjec
    '    oCmdEjec.CommandText = "SP_DELIVERY_DOCTOCLIENTE"
    '    oCmdEjec.CommandType = adCmdStoredProc
    '
    '    'Dim orsD As ADODB.Recordset
    '
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCLIENTE", adBigInt, adParamInput, , frmDeliveryApp.lblCliente.Caption)
    '
    '    Set orsD = oCmdEjec.Execute
    '
    '    Dim Vdocto As String
    '
    '    If Not orsD.EOF Then
    '        Vdocto = Trim(orsD!DOCTO)
    '    End If

    On Error GoTo printe

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim vIgv        As Currency

    Dim vSubTotal   As Currency
    
    'vSubTotal = xSubTotal 'Round((xTotal / ((100 + LK_IGV) / 100)), 2)
    'vSubTotal = Round(((xTotal - xICBPER) / ((100 + LK_IGV) / 100)), 2)

    'vIgv = xIgv ' xTotal - vSubTotal
    'vIgv = Round(vSubTotal * (LK_IGV / 100), 2)

    'If TipoDoc = "B" Then
  
    'Else
    '    oCmdEjec.CommandText = "SpPrintFacDet"
    'End If

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "cliente"

                crParamDef.AddCurrentValue Trim(xcliente)

            Case "FechaEmi"
                crParamDef.AddCurrentValue LK_FECHA_DIA

            Case "Son"
                crParamDef.AddCurrentValue CONVER_LETRAS(CStr(xTotal), "S")

            Case "total"
                crParamDef.AddCurrentValue CStr(FormatNumber(xTotal, 2)) ' CStr(xTotal)

            Case "subtotal"
                crParamDef.AddCurrentValue CStr(FormatNumber(xSubTotal, 2))

            Case "igv"
                crParamDef.AddCurrentValue CStr(FormatNumber(xIgv, 2))

            Case "SerFac"
                crParamDef.AddCurrentValue xSerie

            Case "NumFac"
                crParamDef.AddCurrentValue CStr(xNumero)

            Case "DirClie"

                'crParamDef.AddCurrentValue frmDeliveryApp.DatDireccion.Text
                crParamDef.AddCurrentValue xDireccion

            Case "RucClie"

                'crParamDef.AddCurrentValue Vdocto
                crParamDef.AddCurrentValue xRuc

            Case "Dni" 'linea nueva
                crParamDef.AddCurrentValue xdni 'linea nueva

        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.CommandText = "SpPrintFacturacion"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, xCia)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Serie", adChar, adParamInput, 3, xSerie)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nro", adInteger, adParamInput, , xNumero)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fbg", adChar, adParamInput, 1, Left(xTipoDocto, 1)) ' IIf(Me.ComDocto.ListIndex = 0, "F", IIf(Me.ComDocto.ListIndex = 1, "B", "")))
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroCom", adInteger, adParamInput, , vNroCom)

    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"
    Dim DD As ADODB.Recordset

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
    If Not rsd.EOF Then
        If MsgBox("¿Desea imprimir?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
        VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        
        VReporte.PrintOut False, xNroCopias, , 1, 1
        frmVisor.cr.ReportSource = VReporte
        'frmVisor.cr.ViewReport
        'frmVisor.Show vbModal
        Else
        End If
    
    End If
    
    Set objCrystal = Nothing
    Set VReporte = Nothing
    
    If xEsprom Then
        RutaReporte = PUB_RUTA_REPORTE + "promo.rpt"
        Set VReporte = objCrystal.OpenReport(RutaReporte)
    
        VReporte.PrintOut False, 1, , 1, 1
        frmVisor.cr.ReportSource = VReporte

    End If
    
    Exit Sub

printe:
    MostrarErrores Err

End Sub
