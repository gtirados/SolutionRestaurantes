VERSION 5.00
Begin VB.Form frmAsigna 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese Código"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5040
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblplato 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Código:"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   1395
   End
End
Attribute VB_Name = "frmAsigna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Mostrador As Boolean
Public gDelivery As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Not IsNumeric(Me.txtCantidad.Text) Then
            MsgBox "Debe ingresar solo numeros", vbCritical, "Error"

            Exit Sub

        End If

        If Len(Trim(Me.txtCodigo.Text)) <> 0 Then
            If Not Mostrador Then
 If Not gDelivery Then
                If Len(Trim(frmComanda.lblMozo.Caption)) = 0 Then
                    MsgBox "Debe Elegir Mozo", vbInformation, "Error"

                    Exit Sub

                End If
                End If
            End If
   
            If Mostrador Then
                frmComanda2.oRsPlatos.Filter = "alt='" & Me.txtCodigo.Text & "'"
            Else
                If gDelivery Then
                frmDeliveryApp.oRsPlatos.Filter = "alt='" & Me.txtCodigo.Text & "'"
                Else
                frmComanda.oRsPlatos.Filter = "alt='" & Me.txtCodigo.Text & "'"
                End If
            End If
    
            ' =====ACA OBTENGO EL CODIGO DEL PLATO
            SQ_OPER = 3
            pu_alterno = Me.txtCodigo.Text
            pu_codcia = LK_CODCIA
            LEER_ART_LLAVE

            If art_llave_alt.EOF Then
                MsgBox "Codigo No Existe ...", 48, Pub_Titulo

                Exit Sub

            End If

            PUB_CODART = art_llave_alt!ART_KEY

            Dim oRsTemp As ADODB.Recordset

            'Varificando insumos del plato
            Dim msn     As String

            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpDevuelveInsumosxPlato"
            oCmdEjec.CommandType = adCmdStoredProc
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodPlato", adDouble, adParamInput, , CDbl(PUB_CODART))
            'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@mensaje", adVarChar, adParamOutput, 300, msn)

            Dim vstrmin  As String 'variable para capturar los insumos

            Dim vstrcero As String 'variabla para capturar insumos en cero

            Dim vmin     As Boolean 'minimo

            Dim vcero    As Boolean 'stock minimo

            vmin = False
            vcero = False

            Set oRsTemp = oCmdEjec.Execute

            If Not oRsTemp.EOF Then

                Do While Not oRsTemp.EOF

                    If oRsTemp!sa <= 0 Or (oRsTemp!sa - oRsTemp!ei) < 0 Then
                        vcero = True

                        'MsgBox "Algunos insumos del Plato no estan disponibles", vbCritical, NombreProyecto
                        If Len(vstrcero) = 0 Then
                            vstrcero = Trim(oRsTemp!nm)
                        Else
                            vstrcero = vstrcero & vbCrLf & Trim(oRsTemp!nm)
                        End If

                        'Exit Sub
        
                    ElseIf (oRsTemp!sa - oRsTemp!ei) <= oRsTemp!sm Then
                        vmin = True

                        'MsgBox "Algunos insumos del Plato estan el el Minimó permitido", vbInformation, NombreProyecto
                        If Len(vstrmin) = 0 Then
                            vstrmin = Trim(oRsTemp!nm)
                        Else
                            vstrmin = vstrmin & vbCrLf & Trim(oRsTemp!nm)
                        End If

                        'Exit Do
                    End If

                    'c = c + 1
                    oRsTemp.MoveNext
                Loop

                'Else
                '    MsgBox "El plato no tiene insumos", vbCritical, NombreProyecto
            End If

            If vmin Then
                MsgBox "Los siguientes insumos del Plato estan el el Minimo permitido" & vbCrLf & vstrmin, vbInformation, NombreProyecto
            End If

            If vcero Then
                'MsgBox "Algunos insumos del Plato no estan disponibles" & vbCrLf & vstrcero, vbCritical, NombreProyecto
            End If

            'If vcero Then Exit Sub
            Dim vBusca As Boolean

            Dim f      As Integer

            Dim DD     As Integer

            If Mostrador Then
                If Not frmComanda2.oRsPlatos.EOF Then
                    frmComanda2.vCodFam = frmComanda2.oRsPlatos!codfam

                    If frmComanda2.VNuevo Then
                        If frmComanda2.lvPlatos.ListItems.count = 0 Then
                            If frmComanda2.AgregaPlato(frmComanda2.oRsPlatos!Codigo, Me.txtCantidad.Text, frmComanda2.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmComanda2.oRsPlatos!PRECIO), "", "", 0, frmComanda2.lblCliente.Caption, IIf(Len(Trim(frmComanda2.lblComensales.Caption)) = 0, 0, frmComanda2.lblComensales.Caption)) Then
                                Me.lblplato.Caption = frmComanda2.oRsPlatos!PLATO
                                Set ItemP = frmComanda2.lvPlatos.ListItems.Add(, , Trim(frmComanda2.oRsPlatos!PLATO), frmComanda2.ilComanda.ListImages.Item(1).Key, frmComanda2.ilComanda.ListImages.Item(1).Key)
                                ItemP.Tag = frmComanda2.oRsPlatos!Codigo
                                ItemP.Checked = True
                                ItemP.SubItems(2) = " "
                                ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                ItemP.SubItems(4) = Format(frmComanda2.oRsPlatos!PRECIO, "##.#0")
                                ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                ItemP.SubItems(6) = 0
                                ItemP.SubItems(7) = 0   'linea nueva
                                ItemP.SubItems(8) = frmComanda2.vMaxFac
                                ItemP.SubItems(9) = 0
                                frmComanda2.VNuevo = False
           
                                'ItemP.Checked = True
                                frmComanda2.oRsPlatos.Filter = ""
                                frmComanda2.oRsPlatos.MoveFirst

                                vEstado = "O"
                                Me.txtCodigo.Text = ""
                                Me.txtCantidad.Text = 1
                                Me.txtCodigo.SetFocus
                            End If

                        Else
                    
                            'Dim f As ListItem
                
                            vBusca = False
                
                            For f = 1 To frmComanda2.lvPlatos.ListItems.count

                                If frmComanda2.lvPlatos.ListItems(f).Tag = frmComanda2.oRsPlatos!Codigo Then
                                    vBusca = True

                                    Exit For

                                End If

                            Next

                            If vBusca Then
                                frmComanda2.lvPlatos.ListItems(f).SubItems(3) = FormatNumber(val(frmComanda2.lvPlatos.ListItems(f).SubItems(3)) + 1, 2)
                                frmComanda2.lvPlatos.ListItems(f).SubItems(5) = FormatNumber(val(frmComanda2.lvPlatos.ListItems(f).SubItems(3)) * val(frmComanda2.lvPlatos.ListItems(f).SubItems(4)), 2)
                            Else

                                If frmComanda2.AgregaPlato(frmComanda2.oRsPlatos!Codigo, Me.txtCantidad.Text, frmComanda2.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmComanda2.oRsPlatos!PRECIO), "", "", 0, frmComanda2.lblCliente.Caption, IIf(Len(Trim(frmComanda2.lblComensales.Caption)) = 0, 0, frmComanda2.lblComensales.Caption)) Then
                                    Me.lblplato.Caption = frmComanda2.oRsPlatos!PLATO
                                    Set ItemP = frmComanda2.lvPlatos.ListItems.Add(, , Trim(frmComanda2.oRsPlatos!PLATO), frmComanda2.ilComanda.ListImages.Item(1).Key, frmComanda2.ilComanda.ListImages.Item(1).Key)
                                    ItemP.Tag = frmComanda2.oRsPlatos!Codigo
                                    ItemP.Checked = True
                                    ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                    'obteniendo precio
                                    frmComanda2.oRsPlatos.Filter = "Codigo = '" & frmComanda2.oRsPlatos!Codigo & "'"

                                    If Not frmComanda2.oRsPlatos.EOF Then: ItemP.SubItems(4) = FormatNumber(frmComanda2.oRsPlatos!PRECIO, 2)
                                    ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                    frmComanda2.oRsPlatos.Filter = ""
                                    frmComanda2.oRsPlatos.MoveFirst
                                    Me.txtCodigo.Text = ""
                                    Me.txtCantidad.Text = 1
                                    Me.txtCodigo.SetFocus
                                    Deselecciona

                                End If
                            End If
                        End If

                    Else
            
                        frmComanda2.oRsPlatos.Filter = "Codigo = '" & frmComanda2.oRsPlatos!Codigo & "'"

                        'AgregaPlato Me.cmdPlato(Index).Tag, 1, FormatNumber(frmcomanda2.oRsPlatos!Precio, 2), oRsPlatos!Precio, "", Me.lblSerie.Caption, Me.lblNumero.Caption, dd
                        If frmComanda2.AgregaPlato(frmComanda2.oRsPlatos!Codigo, Me.txtCantidad.Text, frmComanda2.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmComanda2.oRsPlatos!PRECIO), "", frmComanda2.lblSerie.Caption, frmComanda2.lblNumero.Caption, frmComanda2.lblCliente.Caption, IIf(Len(Trim(frmComanda2.lblComensales.Caption)) = 0, 0, frmComanda2.lblComensales.Caption), DD) Then
                            Me.lblplato.Caption = frmComanda2.oRsPlatos!PLATO
                            Set ItemP = frmComanda2.lvPlatos.ListItems.Add(, , Trim(frmComanda2.oRsPlatos!PLATO), frmComanda2.ilComanda.ListImages.Item(1).Key, frmComanda2.ilComanda.ListImages.Item(1).Key)
                            ItemP.Tag = frmComanda2.oRsPlatos!Codigo
                            ItemP.Checked = True
                            ItemP.SubItems(2) = " "
                            ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                            'obteniendo precio
                            frmComanda2.oRsPlatos.Filter = "Codigo = '" & frmComanda2.oRsPlatos!Codigo & "'"

                            If Not frmComanda2.oRsPlatos.EOF Then: ItemP.SubItems(4) = Format(frmComanda2.oRsPlatos!PRECIO, "##.#0")
                            ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                            ItemP.SubItems(6) = DD
                            ItemP.SubItems(7) = 0   'linea nueva
                            ItemP.SubItems(8) = frmComanda2.vMaxFac
                            ItemP.SubItems(9) = 0
                            frmComanda2.oRsPlatos.Filter = ""
                            frmComanda2.oRsPlatos.MoveFirst
                        End If
                    End If

                    frmComanda2.CargarComanda LK_CODCIA, "0"
                    frmComanda2.lblTot.Caption = FormatCurrency(frmComanda2.sumatoria, 2)
                    frmComanda2.lblItems.Caption = "Items: " & frmComanda2.lvPlatos.ListItems.count
                    frmComanda2.lvPlatos.ListItems(frmComanda2.lvPlatos.ListItems.count).Selected = True
                    Me.txtCantidad.Text = 1
                    Me.txtCodigo.Text = ""
                    Me.txtCodigo.SetFocus
                    Deselecciona
                Else
                    MsgBox "Código de Articulo no existe", vbCritical, "Error"
                    Me.txtCodigo.SelStart = 0
                    Me.txtCodigo.SelLength = Len(Me.txtCodigo.Text)
                    Me.txtCodigo.SetFocus
                End If

            Else 'no es mostrador

                If gDelivery Then
                    If Not frmDeliveryApp.oRsPlatos.EOF Then
                        frmDeliveryApp.vCodFam = frmDeliveryApp.oRsPlatos!codfam

                        If frmDeliveryApp.VNuevo Then
                            If frmDeliveryApp.lvPlatos.ListItems.count = 0 Then
                                If frmDeliveryApp.AgregaPlato(frmDeliveryApp.oRsPlatos!Codigo, Me.txtCantidad.Text, frmDeliveryApp.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmDeliveryApp.oRsPlatos!PRECIO), "", "", 0, frmDeliveryApp.lblCliente.Caption, 0) Then
                                    Me.lblplato.Caption = Trim(frmDeliveryApp.oRsPlatos!PLATO)
                                    Set ItemP = frmDeliveryApp.lvPlatos.ListItems.Add(, , Trim(frmDeliveryApp.oRsPlatos!PLATO), frmDeliveryApp.ilPedido.ListImages.Item(1).Key, frmDeliveryApp.ilPedido.ListImages.Item(1).Key)
                                    ItemP.Tag = frmDeliveryApp.oRsPlatos!Codigo
                                    ItemP.Checked = True
                                    ItemP.SubItems(2) = " "
                                    ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                    ItemP.SubItems(4) = Format(frmDeliveryApp.oRsPlatos!PRECIO, "##.#0")
                                    ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                    ItemP.SubItems(6) = 0
                                    ItemP.SubItems(7) = 0   'linea nueva
                                    ItemP.SubItems(8) = frmDeliveryApp.vMaxFac
                                    ItemP.SubItems(9) = 0
                                    frmDeliveryApp.VNuevo = False
           
                                    'ItemP.Checked = True
                                    frmDeliveryApp.oRsPlatos.Filter = ""
                                    frmDeliveryApp.oRsPlatos.MoveFirst

                                    vEstado = "O"
                                   
                                End If

                            Else
                    
                                'Dim f As ListItem
                
                                vBusca = False
                
                                For f = 1 To frmDeliveryApp.lvPlatos.ListItems.count

                                    If frmDeliveryApp.lvPlatos.ListItems(f).Tag = frmDeliveryApp.oRsPlatos!Codigo Then
                                        vBusca = True

                                        Exit For

                                    End If

                                Next

                                If vBusca Then
                                    frmDeliveryApp.lvPlatos.ListItems(f).SubItems(3) = FormatNumber(val(frmDeliveryApp.lvPlatos.ListItems(f).SubItems(3)) + 1, 2)
                                    frmDeliveryApp.lvPlatos.ListItems(f).SubItems(5) = FormatNumber(val(frmDeliveryApp.lvPlatos.ListItems(f).SubItems(3)) * val(frmDeliveryApp.lvPlatos.ListItems(f).SubItems(4)), 2)
                                Else

                                    If frmDeliveryApp.AgregaPlato(frmDeliveryApp.oRsPlatos!Codigo, Me.txtCantidad.Text, frmDeliveryApp.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmDeliveryApp.oRsPlatos!PRECIO), "", "", 0, frmDeliveryApp.lblCliente.Caption, 0) Then
                                        Me.lblplato.Caption = frmDeliveryApp.oRsPlatos!PLATO
                                        Set ItemP = frmDeliveryApp.lvPlatos.ListItems.Add(, , Trim(frmDeliveryApp.oRsPlatos!PLATO), frmDeliveryApp.ilPedido.ListImages.Item(1).Key, frmDeliveryApp.ilPedido.ListImages.Item(1).Key)
                                        ItemP.Tag = frmDeliveryApp.oRsPlatos!Codigo
                                        ItemP.Checked = True
                                        ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                        'obteniendo precio
                                        frmDeliveryApp.oRsPlatos.Filter = "Codigo = '" & frmDeliveryApp.oRsPlatos!Codigo & "'"

                                        If Not frmDeliveryApp.oRsPlatos.EOF Then: ItemP.SubItems(4) = FormatNumber(frmDeliveryApp.oRsPlatos!PRECIO, 2)
                                        ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                        frmDeliveryApp.oRsPlatos.Filter = ""
                                        frmDeliveryApp.oRsPlatos.MoveFirst
                                       Me.txtCodigo.Text = ""
                                    Me.txtCantidad.Text = 1
                                    Me.txtCodigo.SetFocus
                                        Deselecciona

                                    End If
                                End If
                            End If

                        Else
            
                            frmDeliveryApp.oRsPlatos.Filter = "Codigo = '" & frmDeliveryApp.oRsPlatos!Codigo & "'"

                            'AgregaPlato Me.cmdPlato(Index).Tag, 1, FormatNumber(frmcomanda.oRsPlatos!Precio, 2), oRsPlatos!Precio, "", Me.lblSerie.Caption, Me.lblNumero.Caption, dd
                            If frmDeliveryApp.AgregaPlato(frmDeliveryApp.oRsPlatos!Codigo, Me.txtCantidad.Text, frmDeliveryApp.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmDeliveryApp.oRsPlatos!PRECIO), "", frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption, frmDeliveryApp.lblCliente.Caption, 0, DD) Then
                                Me.lblplato.Caption = Trim(frmDeliveryApp.oRsPlatos!PLATO)
                                Set ItemP = frmDeliveryApp.lvPlatos.ListItems.Add(, , Trim(frmDeliveryApp.oRsPlatos!PLATO), frmDeliveryApp.ilPedido.ListImages.Item(1).Key, frmDeliveryApp.ilPedido.ListImages.Item(1).Key)
                                ItemP.Tag = frmDeliveryApp.oRsPlatos!Codigo
                                ItemP.Checked = True
                                ItemP.SubItems(2) = " "
                                ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                'obteniendo precio
                                frmDeliveryApp.oRsPlatos.Filter = "Codigo = '" & frmDeliveryApp.oRsPlatos!Codigo & "'"

                                If Not frmDeliveryApp.oRsPlatos.EOF Then: ItemP.SubItems(4) = Format(frmDeliveryApp.oRsPlatos!PRECIO, "##.#0")
                                ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                ItemP.SubItems(6) = DD
                                ItemP.SubItems(7) = 0   'linea nueva
                                ItemP.SubItems(8) = frmDeliveryApp.vMaxFac
                                ItemP.SubItems(9) = 0
                                frmDeliveryApp.oRsPlatos.Filter = ""
                                frmDeliveryApp.oRsPlatos.MoveFirst
                            End If
                        End If

                        frmDeliveryApp.CargarComanda LK_CODCIA, frmDeliveryApp.lblSerie.Caption, frmDeliveryApp.lblNumero.Caption
                        frmDeliveryApp.lblTot.Caption = FormatCurrency(frmDeliveryApp.sumatoria, 2)
                        frmDeliveryApp.lblItems.Caption = "Items: " & frmDeliveryApp.lvPlatos.ListItems.count
                        frmDeliveryApp.lvPlatos.ListItems(frmDeliveryApp.lvPlatos.ListItems.count).Selected = True
                         Me.txtCodigo.Text = ""
                                    Me.txtCantidad.Text = 1
                                    Me.txtCodigo.SetFocus
                        Deselecciona
                    Else
                        MsgBox "Código de Articulo no existe", vbCritical, "Error"
                        Me.txtCodigo.SelStart = 0
                        Me.txtCodigo.SelLength = Len(Me.txtCodigo.Text)
                        Me.txtCodigo.SetFocus
                    End If

                Else 'no es delivery

                    If Not frmComanda.oRsPlatos.EOF Then
                        frmComanda.vCodFam = frmComanda.oRsPlatos!codfam

                        If frmComanda.VNuevo Then
                            If frmComanda.lvPlatos.ListItems.count = 0 Then
                                If frmComanda.AgregaPlato(frmComanda.oRsPlatos!Codigo, Me.txtCantidad.Text, frmComanda.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmComanda.oRsPlatos!PRECIO), "", "", 0, frmComanda.lblCliente.Caption, IIf(Len(Trim(frmComanda.lblComensales.Caption)) = 0, 0, frmComanda.lblComensales.Caption)) Then
                                    Me.lblplato.Caption = frmComanda.oRsPlatos!PLATO
                                    Set ItemP = frmComanda.lvPlatos.ListItems.Add(, , Trim(frmComanda.oRsPlatos!PLATO), frmComanda.ilComanda.ListImages.Item(1).Key, frmComanda.ilComanda.ListImages.Item(1).Key)
                                    ItemP.Tag = frmComanda.oRsPlatos!Codigo
                                    ItemP.Checked = True
                                    ItemP.SubItems(2) = " "
                                    ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                    ItemP.SubItems(4) = Format(frmComanda.oRsPlatos!PRECIO, "##.#0")
                                    ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                    ItemP.SubItems(6) = 0
                                    ItemP.SubItems(7) = 0   'linea nueva
                                    ItemP.SubItems(8) = frmComanda.vMaxFac
                                    ItemP.SubItems(9) = 0
                                    frmComanda.VNuevo = False
           
                                    'ItemP.Checked = True
                                    frmComanda.oRsPlatos.Filter = ""
                                    frmComanda.oRsPlatos.MoveFirst

                                    For i = 1 To frmDisMesas.imgMesa.count - 1

                                        If frmDisMesas.lblNomMesa(i).Tag = vMesa Then
                                            frmDisMesas.imgMesa(i).Picture = frmDisMesas.ilMesas.ListImages(4).Picture
                                            frmDisMesas.imgMesa(i).ToolTipText = "Mesa Ocupada"
                                            frmDisMesas.imgMesa(i).Tag = "O"
                                        End If

                                    Next

                                    vEstado = "O"
                                    Me.txtCodigo.Text = ""
                                    Me.txtCantidad.Text = 1
                                    Me.txtCodigo.SetFocus
                                End If

                            Else
                    
                                'Dim f As ListItem
                
                                vBusca = False
                
                                For f = 1 To frmComanda.lvPlatos.ListItems.count

                                    If frmComanda.lvPlatos.ListItems(f).Tag = frmComanda.oRsPlatos!Codigo Then
                                        vBusca = True

                                        Exit For

                                    End If

                                Next

                                If vBusca Then
                                    frmComanda.lvPlatos.ListItems(f).SubItems(3) = FormatNumber(val(frmComanda.lvPlatos.ListItems(f).SubItems(3)) + 1, 2)
                                    frmComanda.lvPlatos.ListItems(f).SubItems(5) = FormatNumber(val(frmComanda.lvPlatos.ListItems(f).SubItems(3)) * val(frmComanda.lvPlatos.ListItems(f).SubItems(4)), 2)
                                Else

                                    If frmComanda.AgregaPlato(frmComanda.oRsPlatos!Codigo, Me.txtCantidad.Text, frmComanda.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmComanda.oRsPlatos!PRECIO), "", "", 0, frmComanda.lblCliente.Caption, IIf(Len(Trim(frmComanda.lblComensales.Caption)) = 0, 0, frmComanda.lblComensales.Caption)) Then
                                        Me.lblplato.Caption = frmComanda.oRsPlatos!PLATO
                                        Set ItemP = frmComanda.lvPlatos.ListItems.Add(, , Trim(frmComanda.oRsPlatos!PLATO), frmComanda.ilComanda.ListImages.Item(1).Key, frmComanda.ilComanda.ListImages.Item(1).Key)
                                        ItemP.Tag = frmComanda.oRsPlatos!Codigo
                                        ItemP.Checked = True
                                        ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                        'obteniendo precio
                                        frmComanda.oRsPlatos.Filter = "Codigo = '" & frmComanda.oRsPlatos!Codigo & "'"

                                        If Not frmComanda.oRsPlatos.EOF Then: ItemP.SubItems(4) = FormatNumber(frmComanda.oRsPlatos!PRECIO, 2)
                                        ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                        frmComanda.oRsPlatos.Filter = ""
                                        frmComanda.oRsPlatos.MoveFirst
                                        Me.txtCodigo.Text = ""
                                        Me.txtCantidad.Text = 1
                                        Me.txtCodigo.SetFocus
                                        Deselecciona

                                    End If
                                End If
                            End If

                        Else
            
                            frmComanda.oRsPlatos.Filter = "Codigo = '" & frmComanda.oRsPlatos!Codigo & "'"

                            'AgregaPlato Me.cmdPlato(Index).Tag, 1, FormatNumber(frmcomanda.oRsPlatos!Precio, 2), oRsPlatos!Precio, "", Me.lblSerie.Caption, Me.lblNumero.Caption, dd
                            If frmComanda.AgregaPlato(frmComanda.oRsPlatos!Codigo, Me.txtCantidad.Text, frmComanda.oRsPlatos!PRECIO, CDbl(Me.txtCantidad.Text * frmComanda.oRsPlatos!PRECIO), "", frmComanda.lblSerie.Caption, frmComanda.lblNumero.Caption, frmComanda.lblCliente.Caption, IIf(Len(Trim(frmComanda.lblComensales.Caption)) = 0, 0, frmComanda.lblComensales.Caption), DD) Then
                                Me.lblplato.Caption = frmComanda.oRsPlatos!PLATO
                                Set ItemP = frmComanda.lvPlatos.ListItems.Add(, , Trim(frmComanda.oRsPlatos!PLATO), frmComanda.ilComanda.ListImages.Item(1).Key, frmComanda.ilComanda.ListImages.Item(1).Key)
                                ItemP.Tag = frmComanda.oRsPlatos!Codigo
                                ItemP.Checked = True
                                ItemP.SubItems(2) = " "
                                ItemP.SubItems(3) = Format(Me.txtCantidad.Text, "##.#0")
                                'obteniendo precio
                                frmComanda.oRsPlatos.Filter = "Codigo = '" & frmComanda.oRsPlatos!Codigo & "'"

                                If Not frmComanda.oRsPlatos.EOF Then: ItemP.SubItems(4) = Format(frmComanda.oRsPlatos!PRECIO, "##.#0")
                                ItemP.SubItems(5) = Format(val(ItemP.SubItems(3)) * val(ItemP.SubItems(4)), "##.#0")
                                ItemP.SubItems(6) = DD
                                ItemP.SubItems(7) = 0   'linea nueva
                                ItemP.SubItems(8) = frmComanda.vMaxFac
                                ItemP.SubItems(9) = 0
                                frmComanda.oRsPlatos.Filter = ""
                                frmComanda.oRsPlatos.MoveFirst
                            End If
                        End If

                        frmComanda.CargarComanda LK_CODCIA, frmComanda.vMesa
                        frmComanda.lblTot.Caption = FormatCurrency(frmComanda.sumatoria, 2)
                        frmComanda.lblItems.Caption = "Items: " & frmComanda.lvPlatos.ListItems.count
                        frmComanda.lvPlatos.ListItems(frmComanda.lvPlatos.ListItems.count).Selected = True
                        Me.txtCantidad.Text = 1
                        Me.txtCodigo.Text = ""
                        Me.txtCodigo.SetFocus
                        Deselecciona
                    Else
                        MsgBox "Código de Articulo no existe", vbCritical, "Error"
                        Me.txtCodigo.SelStart = 0
                        Me.txtCodigo.SelLength = Len(Me.txtCodigo.Text)
                        Me.txtCodigo.SetFocus
                    End If
                End If
            End If
        End If
    
    End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.txtCantidad.SelStart = 0
    Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
    Me.txtCantidad.SetFocus
End If
End Sub

Private Sub Deselecciona()

    If Not Mostrador Then
        If gDelivery Then

            For C = 1 To frmDeliveryApp.lvPlatos.ListItems.count
                frmDeliveryApp.lvPlatos.ListItems(C).Selected = False
            Next

            frmDeliveryApp.lvPlatos.ListItems(frmDeliveryApp.lvPlatos.ListItems.count).Selected = True
        Else

            For C = 1 To frmComanda.lvPlatos.ListItems.count
                frmComanda.lvPlatos.ListItems(C).Selected = False
            Next

            frmComanda.lvPlatos.ListItems(frmComanda.lvPlatos.ListItems.count).Selected = True
        End If

    Else

        For C = 1 To frmComanda2.lvPlatos.ListItems.count
            frmComanda2.lvPlatos.ListItems(C).Selected = False
        Next

        frmComanda2.lvPlatos.ListItems(frmComanda2.lvPlatos.ListItems.count).Selected = True
    End If

End Sub
