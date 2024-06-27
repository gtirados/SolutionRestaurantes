VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmContratosAlternativas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alternativas de Contrato"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10305
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
   ScaleHeight     =   7410
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   600
      Left            =   9120
      Picture         =   "frmContratosAlternativas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6720
      Width           =   1110
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "C&onfirmar"
      Height          =   600
      Left            =   8040
      Picture         =   "frmContratosAlternativas.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6720
      Width           =   1110
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContratosAlternativas.frx":0714
            Key             =   "save"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   1058
      ButtonWidth     =   1270
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraAlternativas 
      Caption         =   "Alternativas"
      Height          =   3015
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   10095
      Begin MSDataListLib.DataCombo DatDescuentos 
         Height          =   315
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   7800
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtCantidadA 
         Height          =   285
         Left            =   6000
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Detalles"
         Enabled         =   0   'False
         Height          =   600
         Left            =   8880
         Picture         =   "frmContratosAlternativas.frx":0AAE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   990
      End
      Begin VB.CommandButton cmdQuitarA 
         Caption         =   "&Quitar"
         Height          =   600
         Left            =   8880
         Picture         =   "frmContratosAlternativas.frx":0E38
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1680
         Width           =   990
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   7575
      End
      Begin MSComctlLib.ListView lvAlternativas 
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
         Height          =   195
         Left            =   7200
         TabIndex        =   24
         Top             =   645
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   5160
         TabIndex        =   23
         Top             =   645
         Width           =   840
      End
      Begin VB.Label lblDescuento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   285
         Width           =   1065
      End
   End
   Begin MSComctlLib.ListView lvProductos 
      Height          =   1575
      Left            =   1320
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4245
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame FraDetalle 
      Caption         =   "Detalle"
      Height          =   3015
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   10095
      Begin VB.TextBox txtPrecioPlato 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregarD 
         Caption         =   "&Agregar"
         Height          =   600
         Left            =   8880
         Picture         =   "frmContratosAlternativas.frx":11C2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1080
         Width           =   990
      End
      Begin VB.CommandButton cmdQuitarD 
         Caption         =   "&Quitar"
         Enabled         =   0   'False
         Height          =   600
         Left            =   8880
         Picture         =   "frmContratosAlternativas.frx":154C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   990
      End
      Begin VB.TextBox txtDetalle 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   7575
      End
      Begin MSComctlLib.ListView lvDetalle 
         Height          =   1935
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtCantidad 
         Height          =   285
         Left            =   7800
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
         Height          =   195
         Left            =   7320
         TabIndex        =   21
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Menu mnupopup 
      Caption         =   "ALternativa"
      Visible         =   0   'False
      Begin VB.Menu mnueditar 
         Caption         =   "Editar"
      End
   End
End
Attribute VB_Name = "frmContratosAlternativas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private vModifica As Boolean

Private vBuscar   As Boolean 'variable para la busqueda de Productos

Private loc_key   As Integer

Private vPunto    As Boolean

Public oRSAlt     As ADODB.Recordset

Public oRSComp    As ADODB.Recordset
Private oRSDsctos As ADODB.Recordset
Public vAcepta    As Boolean




Private Sub cmdAgregarD_Click()
    AgregarDetalle
End Sub

Private Sub cmdConfirmar_Click()

    If Me.lvDetalle.ListItems.count = 0 Then Exit Sub
    If MsgBox("¿Seguro desea continuar con la operación.?", vbQuestion + vbYesNo, TituloSistema) = vbYes Then
    
        If Not vModifica Then
            If Len(Trim(Me.txtDescripcion.Text)) = 0 Then
                MsgBox "Debe agregar la Descripción de la alternativa.", vbInformation, TituloSistema
                Me.txtDescripcion.SetFocus

                Exit Sub

            ElseIf Len(Trim(Me.txtCantidadA.Text)) = 0 Then
                MsgBox "Debe ingresar la cantidad de la Alternativa.", vbCritical, TituloSistema
                Me.txtCantidadA.SetFocus

                Exit Sub

            ElseIf val(Me.txtCantidadA.Text) = 0 Then
                MsgBox "La Cantidad de la Alternativa es Incorrecta.", vbCritical, TituloSistema
                Me.txtCantidadA.SetFocus
                Me.txtCantidadA.SelStart = 0
                Me.txtCantidadA.SelLength = Len(Me.txtCantidadA.Text)

                Exit Sub

            ElseIf Len(Trim(Me.txtPrecio.Text)) = 0 Then
                MsgBox "Debe ingresar el Precio de la alternativa.", vbCritical, TituloSistema
                Me.txtPrecio.SetFocus

                Exit Sub

            ElseIf val(Me.txtPrecio.Text) = 0 Then
                MsgBox "EL precio de la Alternativa es incorrecto.", vbCritical, TituloSistema
                Me.txtPrecio.SetFocus

                Exit Sub

            Else

                Dim Item   As Object

                Dim Vnro   As Integer

                Dim vdscto As Double, vVALORDSCTO As Double

                If frmContratos.oRsAlternativas.RecordCount > 0 Then
                    frmContratos.oRsAlternativas.Filter = ""
                    frmContratos.oRsAlternativas.MoveFirst
                End If
        
                Vnro = frmContratos.oRsAlternativas.RecordCount + Me.lvAlternativas.ListItems.count + 1
                Set Item = Me.lvAlternativas.ListItems.Add(, , Trim(Me.txtDescripcion.Text))
       
                Item.Tag = Vnro
                vdscto = 0

                'vdscto = IIf(Len(Trim(Me.txtDescuento.Text)) = 0, 0, Me.txtDescuento.Text)
                If Me.DatDescuentos.BoundText <> "" Then
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SPOBTENERDESCUENTO"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDDSCTO", adDouble, adParamInput, , Me.DatDescuentos.BoundText)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VALOR", adDouble, adParamOutput, , 0)
                    oCmdEjec.Execute
                    vdscto = oCmdEjec.Parameters("@VALOR").Value
                End If

                vVALORDSCTO = ((val(Me.txtCantidadA.Text) * val(Me.txtPrecio.Text)) * vdscto) / 100
                Item.SubItems(1) = Me.DatDescuentos.BoundText
                Item.SubItems(2) = vdscto
                Item.SubItems(3) = Me.txtCantidadA.Text
                Item.SubItems(4) = Me.txtPrecio.Text
                Item.SubItems(5) = (Item.SubItems(3) * Item.SubItems(4))
                Item.SubItems(6) = (Item.SubItems(3) * Item.SubItems(4)) - vVALORDSCTO
                Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).Selected = True
      
                oRSAlt.AddNew
                oRSAlt!CODALTERNATIVA = Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).Tag
                oRSAlt!ALTERNATIVA = Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).Text
                oRSAlt!IDDSCTO = IIf(Len(Trim(Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(1))) = 0, 0, Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(1))
                oRSAlt!DESCUENTO = Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(2)
                oRSAlt!Cantidad = Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(3)
                oRSAlt!PRECIO = Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(4)
                'vdscto = ((Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(3) * Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(4)) * Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(2)) / 100 'PORCENTAJE DESCENTO
                oRSAlt!NETO = (Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(3) * Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(4))
            oRSAlt!BRUTO = (Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(3) * Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(4)) - vVALORDSCTO
                oRSAlt.Update
      
                'Me.FraAlternativas.Enabled = False
                Me.txtDescripcion.Text = ""
                
                Me.txtCantidadA.Text = ""
                Me.txtPrecio.Text = ""
                'Me.lvDetalle.ListItems.Clear
                Me.Toolbar1.Buttons(1).Enabled = False
                'Me.FraDetalle.Enabled = True
                Me.cmdConfirmar.Enabled = True
                Me.cmdLimpiar.Enabled = True
                Me.txtDetalle.SetFocus
            End If
        End If

        Dim vImporte As Double

        vdscto = 0
        vImporte = 0
        vImporte = val(Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(3)) * Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(2)

        If val(Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(1)) <> 0 Then
            vdscto = (vImporte * val(Me.lvAlternativas.ListItems(Me.lvAlternativas.ListItems.count).SubItems(1))) / 100
        End If

        If oRSAlt.RecordCount > 0 Then oRSAlt.MoveFirst

        '    Do While Not oRSAlt.EOF
        '        oRSAlt.Delete
        '        oRSAlt.MoveNext
        '    Loop
        
        Dim fila As Object

        ' For Each Fila In Me.lvAlternativas.ListItems
 
        ' Next
    
        If Not oRSComp.BOF And Not oRSComp.EOF Then

            oRSComp.Filter = "CODALTERNATIVA=" & Me.lvAlternativas.SelectedItem.Tag

            If oRSComp.RecordCount <> 0 Then
                oRSComp.MoveFirst
            End If
         
            Do While Not oRSComp.EOF
                oRSComp.Delete
                oRSComp.MoveNext
            Loop

            'If oRSAlt.RecordCount = 1 Then
            'If Not oRSComp.BOF And Not oRSComp.EOF Then

            For Each fila In Me.lvDetalle.ListItems

                With oRSComp
                    .AddNew
                     .Fields!CODALTERNATIVA = Me.lvAlternativas.SelectedItem.Tag
                    .Fields!CODPLATO = fila.Tag
                    .Fields!Cantidad = fila.SubItems(2)
                    .Fields!PRECIO = fila.SubItems(1)
                    .Fields!PRODUCTO = fila.Text
                    .Fields!Importe = fila.SubItems(3)
                    .Update
                End With

            Next

        Else

            For Each fila In Me.lvDetalle.ListItems

                With oRSComp
                    .AddNew
                    .Fields!CODALTERNATIVA = Me.lvAlternativas.SelectedItem.Tag
                    .Fields!CODPLATO = fila.Tag
                    .Fields!Cantidad = fila.SubItems(2)
                    .Fields!PRECIO = fila.SubItems(1)
                    .Fields!PRODUCTO = fila.Text
                    .Fields!Importe = fila.SubItems(3)
                    .Update
                End With

            Next
    
        End If
   
        Me.cmdLimpiar.Enabled = False
        Me.txtDetalle.Text = ""
        Me.txtCantidad.Text = ""
        Me.lvDetalle.ListItems.Clear
        Me.txtDetalle.Tag = ""
        Me.FraAlternativas.Enabled = True
        Me.FraDetalle.Enabled = False
        Me.txtDescripcion.SetFocus
        Me.lvAlternativas.Enabled = True
        Me.Toolbar1.Buttons(1).Enabled = True
    End If

End Sub

Private Sub cmdLimpiar_Click()
    

    'If Not vModifica Then Me.lvAlternativas.ListItems.Remove Me.lvAlternativas.ListItems.count
    
    Me.Toolbar1.Buttons(1).Enabled = True
    Me.txtCantidad.Text = ""
    Me.txtDetalle.Tag = ""
    Me.txtDescripcion.Text = ""
    Me.DatDescuentos.BoundText = ""
    Me.txtPrecio.Text = ""
    Me.txtCantidadA.Text = ""
    Me.lvAlternativas.Enabled = True
    Me.lvDetalle.ListItems.Clear
    Me.FraAlternativas.Enabled = True
    vModifica = False
    'Me.cmdConfirmar.Enabled = False
End Sub

Private Sub CmdModificar_Click()
    Me.FraAlternativas.Enabled = False
    Me.cmdConfirmar.Enabled = True
    Me.FraDetalle.Enabled = True
    Me.Toolbar1.Buttons(1).Enabled = False
    vModifica = True
    Me.cmdLimpiar.Enabled = True
    
    Me.txtDetalle.SetFocus
End Sub

Private Sub cmdQuitarA_Click()

    If Me.lvDetalle.ListItems.count = 0 Then
        MsgBox "No hay nada para poder eliminar.", vbCritical, TituloSistema

        Exit Sub

    End If

    Dim Item As Integer

    Item = Me.lvAlternativas.SelectedItem.Tag
    
    oRSAlt.MoveFirst
    oRSAlt.Filter = "CODALTERNATIVA=" & Item
    If oRSAlt.RecordCount > 0 Then
    oRSAlt.Delete adAffectCurrent
    Me.lvAlternativas.ListItems.Remove Me.lvAlternativas.SelectedItem.Index
End If
    'ELIMINANDO LA COMPOSICION DE LA ALTERNATIVA
    oRSComp.MoveFirst
    oRSComp.Filter = "CODALTERNATIVA=" & Item

    Do While Not oRSComp.EOF
        oRSComp.Delete
        oRSComp.MoveNext
    Loop

    Me.lvDetalle.ListItems.Clear

    If Me.lvAlternativas.ListItems.count > 0 Then
        oRSAlt.Filter = ""
        oRSAlt.MoveFirst
        oRSComp.Filter = ""
        oRSComp.MoveFirst
    End If

End Sub

Private Sub cmdQuitarD_Click()
    Me.lvDetalle.ListItems.Remove Me.lvDetalle.SelectedItem.Index
    Me.cmdQuitarD.Enabled = False
    CalculoImporte
End Sub

Private Sub DatDescuentos_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Me.txtCantidadA.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    vAcepta = False
    ConfiguraLVs
    vBuscar = True

    Dim Item As Object

    vModifica = False
    Me.FraDetalle.Enabled = False

    If oRSAlt.RecordCount > 0 Then oRSAlt.MoveFirst

    Do While Not oRSAlt.EOF
        Set Item = Me.lvAlternativas.ListItems.Add(, , Trim(oRSAlt!ALTERNATIVA))
        Item.Tag = oRSAlt!CODALTERNATIVA
        Item.SubItems(1) = IIf(IsNull(oRSAlt!IDDSCTO), "", oRSAlt!IDDSCTO)
        Item.SubItems(2) = oRSAlt!DESCUENTO
        Item.SubItems(3) = oRSAlt!Cantidad
        Item.SubItems(4) = oRSAlt!PRECIO
        Item.SubItems(5) = oRSAlt!NETO
        Item.SubItems(6) = oRSAlt!BRUTO
        oRSAlt.MoveNext
    Loop

    If oRSAlt.RecordCount > 0 Then oRSAlt.MoveFirst
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPMAESTRODESCUENTOS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    Set oRSDsctos = oCmdEjec.Execute

    With Me.DatDescuentos
        Set .RowSource = oRSDsctos
        .ListField = "DESCUENTO"
        .BoundColumn = "IDE"
    
    End With
    
End Sub

Private Sub ConfiguraLVs()

    With Me.lvAlternativas
        .ColumnHeaders.Add , , "Alternativa", 4400
        .ColumnHeaders.Add , , "IdDSCTO", 0
        .ColumnHeaders.Add , , "Dscto.", 1000
        .ColumnHeaders.Add , , "Cant.", 1000
        .ColumnHeaders.Add , , "Precio", 1000
        .ColumnHeaders.Add , , "Neto", 1000
        .ColumnHeaders.Add , , "Bruto", 1000
        .Gridlines = True
        .LabelEdit = lvwManual
        .MultiSelect = False
        .FullRowSelect = True
        .View = lvwReport
    End With

    With Me.lvDetalle
        .ColumnHeaders.Add , , "Descripción", 6000
        .ColumnHeaders.Add , , "Precio", 800
        .ColumnHeaders.Add , , "Cant.", 800
        .ColumnHeaders.Add , , "Importe", 900
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
        .MultiSelect = False
        
    End With
    
    With Me.lvProductos
        .ColumnHeaders.Add , , "PRODUCTO", 4000
        .ColumnHeaders.Add , , "PRECIO", 500
        .ColumnHeaders.Add , "CANT.", 800
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
        .MultiSelect = False
    
    End With

End Sub

Private Sub lvAlternativas_ItemClick(ByVal Item As MSComctlLib.ListItem)
If oRSComp.RecordCount > 0 Then oRSComp.MoveFirst

    oRSComp.Filter = "CODALTERNATIVA=" & Me.lvAlternativas.SelectedItem.Tag ' Item.Tag
    vModifica = True
Me.cmdModificar.Enabled = True
Me.cmdQuitarA.Enabled = True
    If Not oRSComp.EOF Then
        Me.lvDetalle.ListItems.Clear

        Dim fila As Object

        Do While Not oRSComp.EOF
            Set fila = Me.lvDetalle.ListItems.Add(, , Trim(oRSComp!PRODUCTO))
            fila.Tag = oRSComp!CODPLATO
            fila.SubItems(1) = oRSComp!PRECIO
            fila.SubItems(2) = oRSComp!Cantidad
            fila.SubItems(3) = oRSComp!Importe
            oRSComp.MoveNext

        Loop

        oRSComp.MoveFirst
    End If

End Sub

Private Sub lvAlternativas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnupopup
End If
End Sub

Private Sub lvDetalle_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.cmdQuitarD.Enabled = True
End Sub

Private Sub mnuEditar_Click()
'With Me.lvAlternativas.SelectedItem
'    frmContratosAlternativaEdit.txtDescripcion.Text = .Text
'    frmContratosAlternativaEdit.DatDsctos.BoundText = .SubItems(1)
'    frmContratosAlternativaEdit.txtCantidad.Text = .SubItems(3)
'    frmContratosAlternativaEdit.txtPrecio.Text = .SubItems(4)
'End With
    frmContratosAlternativaEdit.Show vbModal



End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1

            If Me.lvAlternativas.ListItems.count = 0 Then
                MsgBox "Debe agregar alguna Alternativa.", vbCritical, TituloSistema
                Me.txtDescripcion.SetFocus

                Exit Sub

            End If

            vAcepta = True
            Unload Me
      
    End Select

End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)

    If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then AgregarDetalle
End Sub

Private Sub txtCantidadA_KeyPress(KeyAscii As Integer)

    If SoloNumeros(KeyAscii) Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Me.txtPrecio.SetFocus
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
    If KeyAscii = vbKeyReturn Then Me.DatDescuentos.SetFocus
End Sub







Private Sub txtDetalle_Change()
    vBuscar = True
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > Me.lvProductos.ListItems.count Then loc_key = lvProductos.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvProductos.ListItems.count Then loc_key = lvProductos.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.lvProductos.Visible = False
        Me.txtDetalle.Text = ""
        '        Me.lblDocumento.Caption = ""
        '        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
POSICION:
    lvProductos.ListItems.Item(loc_key).Selected = True
    lvProductos.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    Me.txtDetalle.SelStart = Len(Me.txtDetalle.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtDetalle_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.lvProductos.ListItems.Clear
            LimpiaParametros oCmdEjec
            Dim orsDATA As ADODB.Recordset
            oCmdEjec.CommandText = "SPFILTRAPLATOS"
            Set orsDATA = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtDetalle.Text))

            Dim Item As Object
        
            If Not orsDATA.EOF Then

                Do While Not orsDATA.EOF
                    Set Item = Me.lvProductos.ListItems.Add(, , Trim(orsDATA!PRODUCTO))
                    Item.Tag = Trim(orsDATA!Codigo)
                    Item.SubItems(1) = orsDATA!PRECIO
                    orsDATA.MoveNext
                Loop

                Me.lvProductos.Visible = True
                Me.lvProductos.ListItems(1).Selected = True
                loc_key = 1
                Me.lvProductos.ListItems(1).EnsureVisible
                vBuscar = False
            
                '         If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                '         frmCLI.Show vbModal
                '         End If
            End If

        Else 'ASIGNAR VALORES A LOS CONTROLES

            If Me.txtDetalle.Text <> "" Then
                Me.lvProductos.Visible = False
                Me.txtDetalle.Text = Me.lvProductos.SelectedItem.Text
                Me.txtDetalle.Tag = Me.lvProductos.SelectedItem.Tag
                Me.txtPrecioPlato.Text = Me.lvProductos.SelectedItem.SubItems(1)
                Me.FraDetalle.Enabled = True
                Me.txtCantidad.SetFocus
                Me.txtCantidad.SelStart = 0
                Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
                
            
            End If
            
        End If
    End If

End Sub

Private Sub AgregarDetalle()

    If Me.txtDetalle.Tag = "" Then
        MsgBox "Debe ingresar el Producto.", vbCritical, TituloSistema
        Me.txtDetalle.SetFocus

        Exit Sub

    End If

    If Len(Trim(Me.txtCantidad.Text)) = 0 Then
        MsgBox "Cantidad incorrecta para el Producto.", vbCritical, TituloSistema
        Me.txtCantidad.SetFocus

        Exit Sub

    End If
    
    If val(Me.txtCantidad.Text) = 0 Then
        MsgBox "La Cantidad no puede ser cero (0).", vbCritical, TituloSistema
        Me.txtCantidad.SetFocus

        Exit Sub

    End If
    
    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "La Cantidad debe ser un número.", vbCritical, TituloSistema
        Me.txtCantidad.SetFocus

        Exit Sub

    End If

    'TODO: Add 'AgregarDetalle' body here.
    Dim Item As Object

    If Me.lvDetalle.ListItems.count = 0 Then
        Set Item = Me.lvDetalle.ListItems.Add(, , Trim(Me.txtDetalle.Text))
        Item.Tag = Me.txtDetalle.Tag
        Item.SubItems(1) = Me.txtPrecioPlato.Text
        Item.SubItems(2) = Me.txtCantidad.Text
        Item.SubItems(3) = val(Me.txtCantidad.Text) * val(Me.txtPrecioPlato.Text)
        Me.txtDetalle.Text = ""
        Me.txtCantidad.Text = 1
        Me.txtDetalle.Tag = ""
        Me.txtPrecioPlato.Text = ""
        Me.txtDetalle.SetFocus
    Else

        Dim vENC As Boolean

        vENC = False

        For Each Item In Me.lvDetalle.ListItems

            If Me.txtDetalle.Tag = Item.Tag Then
                vENC = True

                Exit For

            End If

        Next

        If vENC Then
            MsgBox "El Producto elegido ya se encuentra agregado", vbInformation, TituloSistema
                
            Me.txtDetalle.SelStart = 0
            Me.txtDetalle.SelLength = Len(Me.txtDetalle.Text)
            Me.txtDetalle.SetFocus
            'Me.txtDetalle.SelLength = Len(Trim(Me.txtDetalle.Text))
        Else
            Set Item = Me.lvDetalle.ListItems.Add(, , Me.txtDetalle.Text)
            Item.Tag = Me.txtDetalle.Tag
            Item.SubItems(1) = Me.txtPrecioPlato.Text
            Item.SubItems(2) = Me.txtCantidad.Text
            Item.SubItems(3) = val(Me.txtCantidad.Text) * val(Me.txtPrecioPlato.Text)
            Me.txtDetalle.Text = ""
            Me.txtCantidad.Text = 1
            Me.txtPrecioPlato.Text = ""
            Me.txtDetalle.Tag = ""
            Me.txtDetalle.SetFocus
        End If

    End If

    CalculoImporte
End Sub

Private Sub txtPrecio_Change()
  If InStr(txtPrecio.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If
End Sub

Private Sub txtPrecio_GotFocus()

    If InStr(txtPrecio.Text, ".") = 0 Then
        vPunto = False
    Else
        vPunto = True
    End If

End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)

    If NumerosyPunto(KeyAscii) Then KeyAscii = 0
    If Chr(KeyAscii) = "." Then
        If vPunto = True Or Len(Trim(txtPrecio.Text)) = 0 Then
         
            KeyAscii = 0
         
        End If
    End If

    If KeyAscii = vbKeyReturn Then
        If Me.FraDetalle.Enabled = True Then
        cmdConfirmar_Click
        Else
            Me.FraDetalle.Enabled = True
            Me.lvAlternativas.Enabled = False
            Me.txtDetalle.SetFocus
            vModifica = False
            Me.cmdQuitarA.Enabled = False
            Me.lvDetalle.ListItems.Clear
        End If
        
    End If

End Sub

Private Sub CalculoImporte()
    Dim fila As Object
    Dim vImporte As Double, vdscto As Double, vVALOSDSCTO As Double
    vImporte = 0
    For Each fila In Me.lvDetalle.ListItems
        vImporte = vImporte + fila.SubItems(3)
    Next
    
    
    
    If vModifica Then
    If Len(Trim(Me.lvAlternativas.SelectedItem.SubItems(1))) = 0 Or Not IsNumeric(Me.lvAlternativas.SelectedItem.SubItems(1)) Then
        vdscto = 0
    Else
        vdscto = Me.lvAlternativas.SelectedItem.SubItems(1)
    End If
   Me.lvAlternativas.SelectedItem.SubItems(3) = Format(vImporte, "##0.#0")
   vImporte = vImporte * Me.lvAlternativas.SelectedItem.SubItems(2)
    vdscto = (vImporte * vdscto) / 100
    'Me.lvAlternativas.SelectedItem.SubItems(3) = vImporte
    Me.lvAlternativas.SelectedItem.SubItems(4) = Format(vImporte - vdscto, "##0.#0")
    Else
    If Me.DatDescuentos.BoundText = "" Then
     'If Len(Trim(Me.txtDescuento.Text)) = 0 Or Not IsNumeric(Me.txtDescuento.Text) Then
        vdscto = 0
    Else
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPOBTENERDESCUENTO"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDDSCTO", adDouble, adParamInput, , Me.DatDescuentos.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@VALOR", adDouble, adParamOutput, , 0)
    oCmdEjec.Execute
    
        vdscto = oCmdEjec.Parameters("@VALOR").Value
    End If
    vVALOSDSCTO = (vImporte * vdscto) / 100
    Me.txtPrecio.Text = vImporte - vVALOSDSCTO
    End If
End Sub
