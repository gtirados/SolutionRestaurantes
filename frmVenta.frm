VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta Caleta"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
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
   ScaleHeight     =   7290
   ScaleWidth      =   9135
   Begin VB.CommandButton cmdBusca 
      Caption         =   "Buscar"
      Height          =   480
      Left            =   5160
      TabIndex        =   34
      Top             =   6720
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvCliente 
      Height          =   2535
      Left            =   240
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CheckBox chkConsumo 
      Caption         =   "Consumo"
      Height          =   255
      Left            =   7560
      TabIndex        =   28
      Top             =   2040
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      Format          =   160694273
      CurrentDate     =   41975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   480
      Left            =   7800
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   480
      Left            =   6480
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvProducto 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox cboTipoDocto 
      Height          =   315
      ItemData        =   "frmVenta.frx":0000
      Left            =   6360
      List            =   "frmVenta.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtnumero 
      Height          =   285
      Left            =   7320
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtSerie 
      Height          =   285
      Left            =   6360
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin MSComctlLib.ListView lvDatos 
      Height          =   3255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   2475
      Width           =   1455
   End
   Begin VB.TextBox txtProducto 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   7095
   End
   Begin VB.Label lblIgv 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4080
      TabIndex        =   33
      Top             =   6315
      Width           =   1575
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Igv:"
      Height          =   195
      Left            =   3600
      TabIndex        =   32
      Top             =   6360
      Width           =   360
   End
   Begin VB.Label lblSubtotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1080
      TabIndex        =   31
      Top             =   6315
      Width           =   1575
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total:"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label lblDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      TabIndex        =   29
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      Height          =   195
      Left            =   6720
      TabIndex        =   27
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label lblTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7440
      TabIndex        =   26
      Top             =   6315
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      Height          =   195
      Left            =   6360
      TabIndex        =   25
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblproducto 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   7680
      TabIndex        =   24
      Top             =   2160
      Width           =   660
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      Height          =   195
      Left            =   6360
      TabIndex        =   23
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      Height          =   195
      Left            =   7560
      TabIndex        =   22
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Serie"
      Height          =   195
      Left            =   6360
      TabIndex        =   21
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblRUC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   360
      TabIndex        =   20
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblIDcliente 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4920
      TabIndex        =   19
      Top             =   120
      Width           =   555
   End
   Begin VB.Label lblImporte 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6120
      TabIndex        =   18
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label lblPrecio 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cantidad:"
      Height          =   195
      Left            =   2400
      TabIndex        =   15
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:"
      Height          =   195
      Left            =   5160
      TabIndex        =   14
      Top             =   2520
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Precio:"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   825
   End
End
Attribute VB_Name = "frmVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loc_keyC As Integer
Private vBuscarC As Boolean 'variable para la busqueda de clientes
Private loc_keyP As Integer
Private vBuscarP As Boolean 'variable para la busqueda de productos

Private Sub cboTipoDocto_Click()
    If Left(Me.cboTipoDocto.Text, 1) = "B" Then
        Me.lblSubTotal.Visible = False
        Me.lblIgv.Visible = False
        Me.Label11.Visible = False
        Me.Label12.Visible = False
        Me.lblSubTotal.Caption = 0
        Me.lblIgv.Caption = 0
    ElseIf Left(Me.cboTipoDocto.Text, 1) = "F" Then
        Me.lblSubTotal.Visible = True
        Me.lblIgv.Visible = True
        Me.Label11.Visible = True
        Me.Label12.Visible = True
        If IsNumeric(Me.lblTotal.Caption) Then
        Me.lblSubTotal.Caption = Round(CDec(Me.lblTotal.Caption) / CDec((LK_IGV / 100) + 1), 2)
        Else
        Me.lblSubTotal.Caption = "0.00"
        End If
        If IsNumeric(Me.lblTotal.Caption) Then
        Me.lblIgv.Caption = CDec(Me.lblTotal.Caption) - CDec(Me.lblSubTotal.Caption)
        Else
        Me.lblIgv.Caption = "0.00"
        End If
    End If
End Sub

Private Sub cboTipoDocto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.dtpFecha.SetFocus
End Sub

Private Sub cmdBusca_Click()
frmVentaConsulta.Show vbModal
End Sub

Private Sub cmdCancelar_Click()
LimpiaPantalla
End Sub

Private Sub cmdGrabar_Click()

If Len(Trim(Me.lblIDcliente.Caption)) = 0 Then
    MsgBox "Debe ingresar el cliente.", vbInformation, Pub_Titulo
    Me.txtCliente.SetFocus
    Exit Sub
End If
If Len(Trim(Me.txtSerie.Text)) = 0 Then
    MsgBox "Debe ingresar la serie.", vbInformation, Pub_Titulo
    Me.txtSerie.SetFocus
    Exit Sub
End If
If Len(Trim(Me.txtnumero.Text)) = 0 Then
    MsgBox "Debe ingresar el número.", vbInformation, Pub_Titulo
    Me.txtnumero.SetFocus
    Exit Sub
End If
If Me.cboTipoDocto.ListIndex = -1 Then
    MsgBox "Debe elegir el tipo de documento.", vbInformation, Pub_Titulo
    Me.cboTipoDocto.SetFocus
    Exit Sub
End If
If Me.lvDatos.ListItems.count = 0 Then
    MsgBox "Debe ingresar productos.", vbCritical, Pub_Titulo
    Exit Sub
End If
    'validaciones
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_FACTURACION_REGISTRAR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@fecha", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPODOCTO", adChar, adParamInput, 1, Left(Me.cboTipoDocto.Text, 1))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@total", adCurrency, adParamInput, , Me.lblTotal.Caption)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@serie", adChar, adParamInput, 3, Me.txtSerie.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@numero", adBigInt, adParamInput, , Me.txtnumero.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@usuario", adVarChar, adParamInput, 10, LK_CODUSU)
    
    Dim xDET As String, i As Integer
    xDET = ""
    If Me.lvDatos.ListItems.count <> 0 Then
        xDET = "<r>"

        For i = 1 To Me.lvDatos.ListItems.count
            xDET = xDET + "<d "
            xDET = xDET + "idp=""" + Me.lvDatos.ListItems(i).Tag + """ "
            xDET = xDET + "cant=""" & Me.lvDatos.ListItems(i).Text & """ "
            xDET = xDET + "pre=""" & Me.lvDatos.ListItems(i).SubItems(2) & """ "
            xDET = xDET + "imp=""" & Me.lvDatos.ListItems(i).SubItems(3) & """ "
            xDET = xDET + "/>"
        Next

        xDET = xDET + "</r>"
    End If
    
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@detalles", adVarChar, adParamInput, 4000, xDET)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CONSUMO", adBoolean, adParamInput, , Me.chkConsumo.Value)
    If Len(Trim(Me.lblIDcliente.Caption)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@idcliente", adBigInt, adParamInput, , Me.lblIDcliente.Caption)
    End If
            
    Set oRSfp = oCmdEjec.Execute

    If Not oRSfp.EOF Then
        If Split(oRSfp!Mensaje, "=")(0) <> 0 Then 'error
            MsgBox Split(oRSfp!Mensaje, "=")(1), vbCritical, Pub_Titulo
        Else
            MsgBox Split(oRSfp!Mensaje, "=")(1), vbInformation, Pub_Titulo
            Imprimir Left(Me.cboTipoDocto.Text, 1), Me.chkConsumo.Value
            LimpiaPantalla
            Me.txtCliente.SetFocus
        End If
    End If

End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtProducto.SetFocus
End Sub



Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
ConfiguraLV
loc_keyC = -1
loc_keyP = -1
Me.dtpFecha.Value = LK_FECHA_DIA
vBuscarC = True
vBuscarP = True
End Sub

Private Sub lvDatos_DblClick()
frmVentaCantidad.txtCantidad.Text = Me.lvDatos.SelectedItem.Text
frmVentaCantidad.Show vbModal
If frmVentaCantidad.vAcepta Then
    Me.lvDatos.SelectedItem.Text = frmVentaCantidad.gCantidad
    Me.lvDatos.SelectedItem.SubItems(3) = CDec(Me.lvDatos.SelectedItem.Text) * CDec(Me.lvDatos.SelectedItem.SubItems(2))
    calcularTotal
End If
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txtCantidad.Text)) = 0 Then
            MsgBox "Debe ingresar la cantidad.", vbInformation, Pub_Titulo
            Me.txtCantidad.SetFocus
        Else

            Dim itemX As Object
            
            If Me.lvDatos.ListItems.count = 0 Then
                Set itemX = Me.lvDatos.ListItems.Add(, , Me.txtCantidad.Text)
                itemX.SubItems(1) = Trim(Me.txtProducto.Text)
                itemX.SubItems(2) = Me.lblPrecio.Caption
                itemX.SubItems(3) = CDec(Me.lblPrecio.Caption) * CDec(Me.txtCantidad.Text)
                itemX.Tag = Me.lblproducto.Caption
                Me.lblproducto.Caption = ""
                Me.txtProducto.Text = ""
                Me.txtCantidad.Text = ""
                Me.lblPrecio.Caption = ""
                Me.lblImporte.Caption = ""
                calcularTotal
                Me.txtProducto.SetFocus
            Else

                Dim i As Integer, vENC As Boolean

                vENC = False

                For i = 1 To Me.lvDatos.ListItems.count
                    If Me.lvDatos.ListItems(i).Tag = Me.lblproducto.Caption Then
                        vENC = True
                        Exit For
                    End If
                Next

                If vENC Then
                    MsgBox "El item ya se encuentra en la lista.", vbInformation, Pub_Titulo
                Else
                    Set itemX = Me.lvDatos.ListItems.Add(, , Me.txtCantidad.Text)
                    itemX.SubItems(1) = Trim(Me.txtProducto.Text)
                    itemX.SubItems(2) = Me.lblPrecio.Caption
                    itemX.SubItems(3) = CDec(Me.lblPrecio.Caption) * CDec(Me.txtCantidad.Text)
                    itemX.Tag = Me.lblproducto.Caption
                    Me.lblproducto.Caption = ""
                    Me.txtProducto.Text = ""
                    Me.txtCantidad.Text = ""
                    Me.lblPrecio.Caption = ""
                    Me.lblImporte.Caption = ""
                    calcularTotal
                    Me.txtProducto.SetFocus
                End If
            End If

        End If
    End If

End Sub

Private Sub txtCliente_Change()
  vBuscarC = True
  loc_keyC = -1
End Sub

Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_keyC = loc_keyC + 1

        If loc_keyC > Me.lvCliente.ListItems.count Then loc_keyC = lvCliente.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_keyC = loc_keyC - 1

        If loc_keyC < 1 Then loc_keyC = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_keyC = loc_keyC + 17

        If loc_keyC > lvCliente.ListItems.count Then loc_keyC = lvCliente.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_keyC = loc_keyC - 17

        If loc_keyC < 1 Then loc_keyC = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.lvCliente.Visible = False
        'Me.txtDescExterno.Text = ""
        '        Me.lblDocumento.Caption = ""
        '        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
posicion:
    lvCliente.ListItems.Item(loc_keyC).Selected = True
    lvCliente.ListItems.Item(loc_keyC).EnsureVisible
    
    'Me.txtDescExterno.SelStart = Len(Me.txtDescExterno.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscarC Then
            Me.lvCliente.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarCliProv"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adVarChar, adParamInput, 1, "C")
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@parametro", adVarChar, adParamInput, 100, Me.txtCliente.Text)
                
            Dim ORSurb As ADODB.Recordset
                
            Set ORSurb = oCmdEjec.Execute

            Dim Item As Object
        
            If Not ORSurb.EOF Then

                Do While Not ORSurb.EOF
                    Set Item = Me.lvCliente.ListItems.Add(, , ORSurb!Nombre)
                    Item.Tag = ORSurb!CodClie
                    Item.SubItems(1) = ORSurb!RUC
                    Item.SubItems(2) = ORSurb!dir
                    ORSurb.MoveNext
                Loop

                Me.lvCliente.Visible = True
                Me.lvCliente.ListItems(1).Selected = True
                loc_keyC = 1
                Me.lvCliente.ListItems(1).EnsureVisible
                vBuscarC = False
            Else
            End If
        
        Else
'            Me.txtUrb.Text = Me.ListView1.ListItems(loc_key_u).Text
'            Me.lblUrb.Caption = Me.ListView1.ListItems(loc_key_u).Tag

Me.lblIDcliente.Caption = Me.lvCliente.ListItems(loc_keyC).Tag
Me.lblRUC.Caption = Me.lvCliente.ListItems(loc_keyC).SubItems(1)
Me.lblDireccion.Caption = Me.lvCliente.ListItems(loc_keyC).SubItems(2)
Me.txtCliente.Text = Me.lvCliente.ListItems(loc_keyC).Text
'
            Me.lvCliente.Visible = False
            vBuscarC = True
            Me.txtSerie.SetFocus
            

        End If
    End If
End Sub

Private Sub ConfiguraLV()

    With Me.lvCliente
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Cliente", 5000
        .ColumnHeaders.Add , , "Ruc", 0
        .ColumnHeaders.Add , , "dir", 0
  
        .MultiSelect = False
    End With

    With Me.lvProducto
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Producto", 5000
        .ColumnHeaders.Add , , "precio", 1000
  
        .MultiSelect = False
    End With

    With Me.lvDatos
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
  
        .ColumnHeaders.Add , , "Cant"
        .ColumnHeaders.Add , , "Producto", 5000
        .ColumnHeaders.Add , , "Precio"
        .ColumnHeaders.Add , , "importe"
  
        .MultiSelect = False
    End With

End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.cboTipoDocto.SetFocus
End Sub

Private Sub txtProducto_Change()
 vBuscarP = True
  loc_keyP = -1
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_keyP = loc_keyP + 1

        If loc_keyP > Me.lvProducto.ListItems.count Then loc_keyP = lvProducto.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 38 Then
        loc_keyP = loc_keyP - 1

        If loc_keyP < 1 Then loc_keyP = 1
        GoTo posicion
    End If

    If KeyCode = 34 Then
        loc_keyP = loc_keyP + 17

        If loc_keyP > lvProducto.ListItems.count Then loc_keyP = lvProducto.ListItems.count
        GoTo posicion
    End If

    If KeyCode = 33 Then
        loc_keyP = loc_keyP - 17

        If loc_keyP < 1 Then loc_keyP = 1
        GoTo posicion
    End If

    If KeyCode = 27 Then
        Me.lvProducto.Visible = False
        'Me.txtDescExterno.Text = ""
        '        Me.lblDocumento.Caption = ""
        '        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
posicion:
    lvProducto.ListItems.Item(loc_keyP).Selected = True
    lvProducto.ListItems.Item(loc_keyP).EnsureVisible
    
    'Me.txtDescExterno.SelStart = Len(Me.txtDescExterno.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        If vBuscarP Then
            Me.lvProducto.ListItems.Clear
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SPPRODUCTOS_SEARCH"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 80, Me.txtProducto.Text)
            
                
            Dim ORSurb As ADODB.Recordset
                
            Set ORSurb = oCmdEjec.Execute

            Dim Item As Object
        
            If Not ORSurb.EOF Then

                Do While Not ORSurb.EOF
                    Set Item = Me.lvProducto.ListItems.Add(, , ORSurb!plato)
                    Item.Tag = ORSurb!Codigo
                    Item.SubItems(1) = ORSurb!PRECIO
                    ORSurb.MoveNext
                Loop

                Me.lvProducto.Visible = True
                Me.lvProducto.ListItems(1).Selected = True
                loc_keyP = 1
                Me.lvProducto.ListItems(1).EnsureVisible
                vBuscarP = False
            Else

            End If
        
        Else

            Me.lblproducto.Caption = Me.lvProducto.ListItems(loc_keyP).Tag
            Me.lblPrecio.Caption = Me.lvProducto.ListItems(loc_keyP).SubItems(1)
            Me.txtProducto.Text = Me.lvProducto.ListItems(loc_keyP).Text
            Me.lvProducto.Visible = False
            vBuscarP = True
            Me.txtCantidad.SetFocus

        End If
    End If

End Sub

Private Sub calcularTotal()

    Dim Item As Object, cTot As Double

    cTot = 0

    For Each Item In Me.lvDatos.ListItems

        cTot = cTot + Item.SubItems(3)
    
    Next

    Me.lblTotal.Caption = cTot

    If Left(Me.cboTipoDocto.Text, 1) = "B" Then
        Me.lblIgv.Caption = "0.00"
        Me.lblSubTotal.Caption = "0.00"
    ElseIf Left(Me.cboTipoDocto.Text, 1) = "F" Then
        Me.lblSubTotal.Caption = Round(CDec(Me.lblTotal.Caption) / CDec((LK_IGV / 100) + 1), 2)
        Me.lblIgv.Caption = CDec(Me.lblTotal.Caption) - CDec(Me.lblSubTotal.Caption)
    End If

End Sub

Private Sub LimpiaPantalla()
Me.lvDatos.ListItems.Clear
Me.txtProducto.Text = ""
Me.lblPrecio.Caption = ""
Me.txtCantidad.Text = ""
Me.lblImporte.Caption = ""
Me.txtCliente.Text = ""
Me.lblRUC.Caption = ""
Me.cboTipoDocto.ListIndex = -1
Me.txtSerie.Text = ""
Me.lblDireccion.Caption = ""
Me.txtnumero.Text = ""
Me.lblTotal.Caption = "0.00"
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.txtnumero.SetFocus
    Me.txtnumero.SelStart = 0
    Me.txtnumero.SelLength = Len(Me.txtnumero.Text)
End If
End Sub


Private Sub Imprimir(TipoDoc As String, Esconsumo As Boolean)

    On Error GoTo printe

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.APPLICATION

    Dim vIgv        As Currency

    Dim vSubTotal   As Currency

    Dim RutaReporte As String

    If Esconsumo Then
        If TipoDoc = "B" Then
            RutaReporte = "C:\Admin\Nordi\BolCon.rpt"
        ElseIf TipoDoc = "F" Then
            RutaReporte = "C:\Admin\Nordi\FacCon.rpt"
            vSubTotal = Me.lblSubTotal.Caption ' Round((Me.lblImporte.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lblIgv.Caption ' Me.lblTotal.Caption - vSubTotal
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        
        End If

    Else

        If TipoDoc = "B" Then
            RutaReporte = "C:\Admin\Nordi\BolDet.rpt"
        ElseIf TipoDoc = "F" Then
            RutaReporte = "C:\Admin\Nordi\FacDet.rpt"
            vSubTotal = Me.lblSubTotal.Caption 'Round((Me.lblSubtotal.Caption / ((100 + LK_IGV) / 100)), 2)
            'vSubTotal = Round((Me.lblImporte.Caption / ((100 + LK_IGV + 5) / 100)), 2)
            'vrec = Round(vSubTotal * 0.05, 2)
            vIgv = Me.lblIgv.Caption
            'vIgv = Me.lblImporte.Caption - vSubTotal - vrec
        End If
    End If


    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "cliente"
                'crParamDef.AddCurrentValue IIf(Len(Trim(Me.txtRS.Text)) = 0, "CLIENTES VARIOS", Trim(Me.txtRS.Text))
                crParamDef.AddCurrentValue Me.txtCliente.Text

            Case "FechaEmi"
                crParamDef.AddCurrentValue LK_FECHA_DIA

            Case "Son"
                crParamDef.AddCurrentValue CONVER_LETRAS(Me.lblTotal.Caption, "S")

            Case "total"
                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)

            Case "subtotal"
                crParamDef.AddCurrentValue CStr(vSubTotal)

            Case "igv"
                crParamDef.AddCurrentValue CStr(vIgv)

            Case "SerFac"
                crParamDef.AddCurrentValue Me.txtSerie.Text

            Case "NumFac"
                crParamDef.AddCurrentValue Me.txtnumero.Text

            Case "DirClie"
                crParamDef.AddCurrentValue Me.lblDireccion.Caption

            Case "RucClie"
                crParamDef.AddCurrentValue Me.lblRUC.Caption

            Case "Importe" 'linea nueva
                crParamDef.AddCurrentValue FormatNumber(Me.lblTotal.Caption, 2)  'linea nueva
                ' Case "rec"          'SR BEFE
                '     crParamDef.AddCurrentValue CStr(vrec)

        End Select

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc

    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.CommandText = "SP_FACTURACION_PRINT"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Serie", adChar, adParamInput, 3, Me.txtSerie.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@nUMERO", adBigInt, adParamInput, , Me.txtnumero.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPO", adChar, adParamInput, 1, Left(Me.cboTipoDocto.Text, 1))
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NroCom", adInteger, adParamInput, , vNroCom)

    Set rsd = oCmdEjec.Execute

    'COCINA
    'rsd.Filter = "PED_FAMILIA=2"
    Dim DD As ADODB.Recordset

    ' For i = 0 To Printers.count - 1
    '        MsgBox Printers(i).DeviceName
    '    Next
    If Not rsd.EOF Then

        VReporte.DataBase.SetDataSource rsd, 3, 1 'lleno el objeto reporte
        'VReporte.SelectPrinter Printer.DriverName, "\\laptop\doPDF v6", Printer.Port
        '
        VReporte.PrintOut False, 1, , 1, 1
'        frmVisor.cr.ReportSource = VReporte
        'frmVisor.cr.ViewReport
        'frmVisor.Show vbModal
    
    End If


    Set objCrystal = Nothing
    Set VReporte = Nothing

    Exit Sub

printe:
    MostrarErrores Err

End Sub
