VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRequerimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Requerimiento"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9210
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   480
      Left            =   3960
      TabIndex        =   12
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   480
      Left            =   6840
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   480
      Left            =   5400
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Del"
      Height          =   360
      Left            =   8280
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   360
      Left            =   8280
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   7080
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvProducto 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtProducto 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin MSComctlLib.ListView lvRequerimiento 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblCosto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5760
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Costo:"
      Height          =   195
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   195
      Left            =   7080
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmRequerimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_key  As Integer
Private vBuscar As Boolean 'variable para la busqueda de clientes

Private Sub cmdAdd_Click()

    If Len(Trim(Me.txtProducto.Text)) = 0 Then
        MsgBox "Ingrese el producto", vbInformation, Pub_Titulo
        Me.txtProducto.SetFocus

        Exit Sub

    End If

  If Len(Trim(Me.lblCosto.Caption)) = 0 Then
        MsgBox "Ingrese el producto", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus

        Exit Sub
    End If
    
    If Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Ingrese la cantidad", vbInformation, Pub_Titulo
        Me.txtCantidad.SetFocus

        Exit Sub

    End If
    
    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "la cantidad es incorrecta.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Exit Sub
    End If

    Dim itemP As Object

    If Me.lvRequerimiento.ListItems.count = 0 Then
        Set itemP = Me.lvRequerimiento.ListItems.Add(, , Me.lblCosto.Tag)
        itemP.SubItems(1) = Me.txtProducto.Text
        itemP.SubItems(2) = Me.txtCantidad.Text
        itemP.SubItems(3) = Me.lblCosto.Caption
    Else

        Dim vENC As Boolean

        vENC = False

        For Each itemP In Me.lvRequerimiento.ListItems

            If itemP.Text = Me.lblCosto.Tag Then
                vENC = True

                Exit For

            End If

        Next

        If Not vENC Then
            Set itemP = Me.lvRequerimiento.ListItems.Add(, , Me.lblCosto.Tag)
            itemP.SubItems(1) = Me.txtProducto.Text
            itemP.SubItems(2) = Me.txtCantidad.Text
            itemP.SubItems(3) = Me.lblCosto.Caption
        Else
            MsgBox "Dato repetido", vbCritical, Pub_Titulo

            Exit Sub

        End If
    End If

    Me.lblCosto.Caption = ""
    Me.txtProducto.Text = ""
    Me.txtCantidad.Text = ""
    Me.txtProducto.SetFocus
End Sub

Private Sub cmdCancelar_Click()
 Me.lvRequerimiento.ListItems.Clear
    Me.lblCosto.Caption = ""
    Me.txtCantidad.Text = ""
    Me.txtProducto.Text = ""
    Me.lvProducto.Visible = False
    Me.txtProducto.SetFocus
End Sub

Private Sub cmdDel_Click()

    If Me.lvRequerimiento.ListItems.count = 0 Then Exit Sub
    Me.lvRequerimiento.ListItems.Remove Me.lvRequerimiento.SelectedItem.Index
End Sub

Private Sub cmdGrabar_Click()
If Me.lvRequerimiento.ListItems.count = 0 Then
    MsgBox "Debe agregar un producto", vbCritical, Pub_Titulo
    Exit Sub
End If
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REQUERIMIENTO_INSERT"

    Dim Xidr As Double

    Xidr = 0
    Pub_ConnAdo.BeginTrans

    On Error GoTo Almacena

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 20, LK_CODUSU)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREQ", adBigInt, adParamOutput, , Xidr)
    oCmdEjec.Execute
    Xidr = oCmdEjec.Parameters("@IDREQ").Value
    
    oCmdEjec.CommandText = "SP_REQUERIMIENTO_DETALLE_INSERT"

    Dim ITEMr As Object

    For Each ITEMr In Me.lvRequerimiento.ListItems

        LimpiaParametros oCmdEjec
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREQUERIMIENTO", adBigInt, adParamInput, , Xidr)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , ITEMr.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COSTO", adDouble, adParamInput, , ITEMr.SubItems(3))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, , ITEMr.SubItems(2))
        oCmdEjec.Execute
    Next

    Pub_ConnAdo.CommitTrans
    Me.lvRequerimiento.ListItems.Clear
    Me.lblCosto.Caption = ""
    Me.txtCantidad.Text = ""
    Me.txtProducto.Text = ""
    Me.lvProducto.Visible = False
    Me.txtProducto.SetFocus
    MsgBox "Datos almacenados correctamente.", vbInformation, Pub_Titulo
    Visualizar Xidr

    Exit Sub

Almacena:
    Pub_ConnAdo.RollbackTrans
    MsgBox Err.Description, vbCritical, Pub_Titulo
                    
End Sub

Private Sub cmdSearch_Click()
    frmRequerimientoSearch.Show vbModal
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
vBuscar = True
    With Me.lvProducto
        .Gridlines = True
        .LabelEdit = lvwAutomatic
        .View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "", 4000
        .ColumnHeaders.Add , , "", 0
        .ColumnHeaders.Add , , "", 500
    End With

    With Me.lvRequerimiento
        .Gridlines = True
        .LabelEdit = lvwAutomatic
        .View = lvwReport
        .ColumnHeaders.Add , , "codigo"
        .ColumnHeaders.Add , , "producto"
        .ColumnHeaders.Add , , "cantidad"
        .ColumnHeaders.Add , , "costo"
        .HideColumnHeaders = False
    End With

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAdd_Click
End Sub

Private Sub txtProducto_Change()
vBuscar = True
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > Me.lvProducto.ListItems.count Then loc_key = lvProducto.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvProducto.ListItems.count Then loc_key = lvProducto.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.lvProducto.Visible = False
        'Me.txtRS.Text = ""
       ' Me.txtRuc.Text = ""
       ' Me.txtDireccion.Text = ""
    End If

    GoTo fin
posicion:
    lvProducto.ListItems.Item(loc_key).Selected = True
    lvProducto.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
   ' txtRS.SelStart = Len(txtRS.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    Dim orsPro As ADODB.Recordset

    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.lvProducto.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SPPRODUCTOS_SEARCH1"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@search", adVarChar, adParamInput, 80, Trim(Me.txtProducto.Text))
        
            Set orsPro = oCmdEjec.Execute

            Dim Item As Object
        
            If Not orsPro.EOF Then

                Do While Not orsPro.EOF
                    Set Item = Me.lvProducto.ListItems.Add(, , orsPro!plato)
                    Item.Tag = orsPro!Codigo
                    Item.SubItems(1) = Trim(orsPro!costo)
                    Item.SubItems(2) = Trim(orsPro!stock)
                    orsPro.MoveNext
                Loop

                Me.lvProducto.Visible = True
                Me.lvProducto.ListItems(1).Selected = True
                loc_key = 1
                Me.lvProducto.ListItems(1).EnsureVisible
                vBuscar = False
            Else

                MsgBox "El producto no existe", vbInformation, Pub_Titulo
            End If
        
        Else
            Me.lblCosto.Caption = Me.lvProducto.ListItems(loc_key).SubItems(1)
            Me.txtProducto.Text = Me.lvProducto.ListItems(loc_key).Text
            Me.lblCosto.Tag = Me.lvProducto.ListItems(loc_key).Tag
        
            Me.lvProducto.Visible = False
            Me.txtCantidad.SetFocus
            'Me.lvDetalle.SetFocus
        End If
    End If

End Sub

Private Sub ConfiguraLV()
With Me.lvProducto
    .FullRowSelect = True
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .ColumnHeaders.Add , , "Codigo", 1000
    .ColumnHeaders.Add , , "Cliente", 5000
    .ColumnHeaders.Add , , "STOCK", 500
    .ColumnHeaders.Add , , "Direcion", 0
    .MultiSelect = False
End With
End Sub

Private Sub Visualizar(Dnro As Double)
  LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REQUERIMIENTO_VIEW"
    oCmdEjec.CommandType = adCmdStoredProc
    
    
    Dim RutaReporte As String


    RutaReporte = PUB_RUTA_REPORTE & "Req.rpt"
    

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

'    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
'
'    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition
'
   Dim objCrystal  As New CRAXDRT.APPLICATION

'    Dim vIgv        As Currency
'
'    Dim vSubTotal   As Currency
'
'
'vSubTotal = Round((xTotal / ((100 + LK_IGV) / 100)), 2)
'
'vIgv = xTotal - vSubTotal

  

    

    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
'    Set crParamDefs = VReporte.ParameterFields
'
'    For Each crParamDef In crParamDefs
'
'        Select Case crParamDef.ParameterFieldName
'
'            Case "cliente"
'
'                crParamDef.AddCurrentValue Trim(xcliente)
'
'            Case "FechaEmi"
'                crParamDef.AddCurrentValue LK_FECHA_DIA
'
'            Case "Son"
'                crParamDef.AddCurrentValue CONVER_LETRAS(CStr(xTotal), "S")
'
'            Case "total"
'                crParamDef.AddCurrentValue CStr(FormatNumber(xTotal, 2)) ' CStr(xTotal)
'
'            Case "subtotal"
'                crParamDef.AddCurrentValue CStr(FormatNumber(vSubTotal, 2))
'
'            Case "igv"
'                crParamDef.AddCurrentValue CStr(FormatNumber(vIgv, 2))
'
'            Case "SerFac"
'                crParamDef.AddCurrentValue XsERIE
'
'            Case "NumFac"
'                crParamDef.AddCurrentValue CStr(xNumero)
'
'            Case "DirClie"
'
'                'crParamDef.AddCurrentValue frmDeliveryApp.DatDireccion.Text
'                crParamDef.AddCurrentValue xDireccion
'
'            Case "RucClie"
'
'                'crParamDef.AddCurrentValue Vdocto
'                crParamDef.AddCurrentValue xRuc
'
'            'Case "Importe" 'linea nueva
'                'crParamDef.AddCurrentValue frmDeliveryApp.lblTot.Caption 'linea nueva
'
'
'        End Select
'
'    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nroreq", adBigInt, adParamInput, , Dnro)
    
    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"
  

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
    If Not rsd.EOF Then

        VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '
        'VReporte.PrintOut False, 1, , 1, 1
        frmVisor.cr.ReportSource = VReporte
        frmVisor.cr.ViewReport
        frmVisor.Show vbModal
    
    End If

    'Set VReporte = Nothing
    'Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    'cr.DataSource = VReporte
    'cr.Destination = crptToWindow
    '
    'rsd.Filter = "PED_FAMILIA=3"
    'If Not rsd.EOF Then
    '    VReporte.Database.SetDataSource rsd, 3, 1 'lleno el objeto reporte
    '    VReporte.SelectPrinter Printer.DriverName, "\\SERVIDOR\Canon MP140 series Printer", Printer.Port 'doPDF v6
    '    VReporte.PrintOut ' , 1, , 1, 1
    'End If

    Set objCrystal = Nothing
    Set VReporte = Nothing

    Exit Sub

printe:
    MostrarErrores Err

End Sub
