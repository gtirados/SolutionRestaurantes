Attribute VB_Name = "modQR"
Public wdsn As String
Public wAcceso As String
Public oRSmain As ADODB.Recordset
Public oCMDmain As ADODB.Command
Public objConex As ADODB.Connection
Public xCadena As String

Public Enum TErrorCorretion
    QualityLow
    QualityMedium
    QualityStandard
    QualityHigh
End Enum
 
 'Lib "D:\Aplicaciones Clientes\SFS\SFS 1.2\SolutionRestaurantesSFSv1.2\quricol32.dll"
Public Declare Sub GenerateBMP _
                Lib "quricol32.dll" _
                Alias "GenerateBMPW" ( _
                ByVal FileName As Long, _
                ByVal Text As Long, _
                ByVal Margin As Long, _
                ByVal Size As Long, _
                ByVal Level As TErrorCorretion)

Public Sub CreaCodigoQR(cTipoDocto As String, _
                         cTipoDoctoVenta As String, _
                         cSerie As String, _
                         cNumero As String, _
                         cFecha As Date, _
                         cIgv As String, _
                         cTotal As String, _
                         cRuc As String, cDni As String)

    Dim xRuc              As String

    Dim XTipoDoctoCliente As String

    Dim xNroDoctoCliente  As String

    Dim xCadena           As String
    Dim xCadenaCNN           As String
    
    If Len(Trim(cRuc)) = 0 Then
        XTipoDoctoCliente = "1"
        xNroDoctoCliente = IIf(Len(Trim(cDni)) = 0, "11111111", cDni)
    Else
        XTipoDoctoCliente = "6"
        xNroDoctoCliente = cRuc
    End If

    Dim xmes       As String, xDia As String

    Dim xTDocventa As String

    xRuc = Leer_Ini(App.Path & "\config.ini", "RUC", "")

    If cTipoDoctoVenta = "01" Then
        xTDocventa = "F"
    ElseIf cTipoDoctoVenta = "03" Then
        xTDocventa = "B"
    End If
    
    xmes = Right("00" + CStr(Month(cFecha)), 2)
    xDia = Right("00" + CStr(Day(cFecha)), 2)

    xCadena = cTipoDocto + "|" + xRuc + "|" + XTipoDoctoCliente + "|" + xNroDoctoCliente + "|" + cTipoDoctoVenta + "|" + xTDocventa + Right("000" + Trim(cSerie), 3) + "-" + cNumero + "|" + CStr(Year(cFecha)) + "-" + xmes + "-" + xDia + "|" + cIgv + "|" + cTotal
                    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "USP_GENERA_DATOSQR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPODOCUMENTO", adChar, adParamInput, 2, cTipoDoctoVenta)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SerDoc", adChar, adParamInput, 4, xTDocventa + Right("000" + Trim(cSerie), 3))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroDoc", adDouble, adParamInput, , cNumero)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , cFecha)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NSERIE", adVarChar, adParamInput, 3, Trim(cSerie))
                
    Dim oRSresul As ADODB.Recordset
    
    Dim objConex As ADODB.Connection
    
    xCadenaCNN = "dsn=" + wdsn + ";uid=sa;pwd=" & claveAcceso & ";database=bdatos;"
    Set objConex = New ADODB.Connection
    objConex.CursorLocation = adUseClient
    objConex.Open xCadenaCNN

    Set oRSresul = oCmdEjec.Execute

    If Not oRSresul.EOF Then
        
        If oRSresul!Exito = 0 Then

            GenerateBMP StrPtr(App.Path & "\codigoQR.jpg"), StrPtr(xCadena), 1, 5, QualityStandard

            Dim oRS As ADODB.Recordset

            Set oRS = New ADODB.Recordset
            oRS.Open "select codigoqr from documentos_qr where codcia='" + LK_CODCIA + "' and TIPODOCUMENTO = '" + cTipoDoctoVenta + "' AND SERIE ='" + xTDocventa + Right("000" + Trim(cSerie), 3) + "'" + " AND NUMERO =" + cNumero, objConex, adOpenKeyset, adLockOptimistic

            Set m_stream = New ADODB.Stream
            m_stream.Type = adTypeBinary
            m_stream.Open
            m_stream.LoadFromFile App.Path & "\codigoQR.jpg"
            oRS.Fields("codigoqr").Value = m_stream.Read
            oRS.Update

            oRS.Close
            objConex.Close
            'cn.Close
        End If
    End If

End Sub

