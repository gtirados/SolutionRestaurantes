VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Begin VB.Form frmListaProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catalogo de Productos"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15705
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15705
   Begin ComctlLib.Toolbar tbBotones 
      Height          =   630
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   1111
      ButtonWidth     =   1376
      ButtonHeight    =   1005
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   3
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Nuevo"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Caption         =   "&Modificar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Eliminar"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraData 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   15495
      Begin MSComctlLib.ListView lvData 
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   15255
         _ExtentX        =   26908
         _ExtentY        =   10186
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Criterio de Busqueda de Información"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   15495
      Begin MSDataListLib.DataCombo DatFamilia 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   360
         Width           =   9015
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   3360
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox cboTipoProd 
         Height          =   315
         ItemData        =   "frmListaProd.frx":0000
         Left            =   120
         List            =   "frmListaProd.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo datSubFamilia 
         Height          =   315
         Left            =   6240
         TabIndex        =   10
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubFamilia:"
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Familia:"
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   5160
         TabIndex        =   8
         Top             =   405
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   2640
         TabIndex        =   7
         Top             =   405
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmListaProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oRsProd As ADODB.Recordset
Private vTIPO As String
Private VNuevo As Boolean
'P = Plato
'M = Insumo
'C = Combo

Private Sub ConfigurarLV()

    With Me.lvData
        .ColumnHeaders.Add , , "Codigo", 1000
        .ColumnHeaders.Add , , "Descripción del Producto", 4000
        .ColumnHeaders.Add , , "Und"
        .ColumnHeaders.Add , , "Linea"
        .ColumnHeaders.Add , , "Lista 1", 800
        .ColumnHeaders.Add , , "Lista 2", 800
        .ColumnHeaders.Add , , "Lista 3", 800
        .ColumnHeaders.Add , , "Lista 4", 800
        .ColumnHeaders.Add , , "Lista 5", 800
        .ColumnHeaders.Add , , "Lista 6", 800
        .ColumnHeaders.Add , , "Stock", 0
        .ColumnHeaders.Add , , "IdFam", 0
        .ColumnHeaders.Add , , "IdSumFam", 0
        .ColumnHeaders.Add , , "Stock", 1000
        .ColumnHeaders.Add , , "Formula", 1000
        .Gridlines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .View = lvwReport
    End With

End Sub

Private Sub CargarFamilias()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPRODUCTOS_LISTARFAMILIAS"

    Dim ORSf As ADODB.Recordset

    Set ORSf = oCmdEjec.Execute(, LK_CODCIA)

    Set Me.DatFamilia.RowSource = ORSf
    Me.DatFamilia.BoundColumn = ORSf.Fields(0).Name
    Me.DatFamilia.ListField = ORSf.Fields(1).Name
    Me.DatFamilia.BoundText = "-1"
End Sub

Private Sub CargarDatos()

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpListarPlatos"

    With oCmdEjec
        .Parameters.Append .CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        .Parameters.Append .CreateParameter("@Tipo", adChar, adParamInput, 1, vTIPO)
        .Parameters.Append .CreateParameter("@FAMILIA", adInteger, adParamInput, , Me.DatFamilia.BoundText)
        .Parameters.Append .CreateParameter("@SUBFAMILIA", adInteger, adParamInput, , Me.datSubFamilia.BoundText)
    End With

    Set oRsProd = oCmdEjec.Execute

    Dim itemX As Object

    Me.lvData.ListItems.Clear

    Do While Not oRsProd.EOF
        Set itemX = Me.lvData.ListItems.Add(, , oRsProd!Codigo)
        itemX.SubItems(1) = Trim(oRsProd!producto)
        itemX.SubItems(2) = Trim(oRsProd!UNIDAD)
        itemX.SubItems(3) = IIf(IsNull(oRsProd!Familia), "", Trim(oRsProd!Familia))
        itemX.SubItems(4) = oRsProd.Fields(4)
        itemX.SubItems(5) = oRsProd.Fields(5)
        itemX.SubItems(6) = oRsProd.Fields(6)
        itemX.SubItems(7) = oRsProd.Fields(7)
        itemX.SubItems(8) = oRsProd.Fields(8)
        itemX.SubItems(9) = oRsProd.Fields(9)
        itemX.SubItems(10) = oRsProd!stock
        itemX.SubItems(11) = oRsProd!idfam
        itemX.SubItems(12) = oRsProd!idsubfam
        itemX.Checked = IIf(oRsProd!SIT, 0, 1)
        itemX.SubItems(13) = oRsProd!descontar
        itemX.SubItems(14) = oRsProd!formulacion

        If oRsProd!SIT = 1 Then
            itemX.ForeColor = vbRed
            itemX.ListSubItems(1).ForeColor = vbRed
            itemX.ListSubItems(2).ForeColor = vbRed
            itemX.ListSubItems(3).ForeColor = vbRed
            itemX.ListSubItems(4).ForeColor = vbRed
            itemX.ListSubItems(5).ForeColor = vbRed
            itemX.ListSubItems(6).ForeColor = vbRed
            itemX.ListSubItems(7).ForeColor = vbRed
            itemX.ListSubItems(8).ForeColor = vbRed
            itemX.ListSubItems(9).ForeColor = vbRed
            itemX.ListSubItems(10).ForeColor = vbRed
            itemX.ListSubItems(11).ForeColor = vbRed
            itemX.ListSubItems(12).ForeColor = vbRed
            itemX.ListSubItems(13).ForeColor = vbRed
            itemX.ListSubItems(14).ForeColor = vbRed
        End If

        oRsProd.MoveNext
    Loop

End Sub


Private Sub cboTipoProd_Change()
If Me.cboTipoProd.ListIndex = 0 Then 'Producto
    vTIPO = "P"
ElseIf Me.cboTipoProd.ListIndex = 1 Then 'Insumo
    vTIPO = "M"
ElseIf Me.cboTipoProd.ListIndex = 2 Then 'Combo
    vTIPO = "C"
End If
CargarDatos
End Sub

Private Sub cboTipoProd_Click()

    If Me.cboTipoProd.ListIndex = 0 Then 'Producto
        vTIPO = "P"
    ElseIf Me.cboTipoProd.ListIndex = 1 Then 'Insumo
        vTIPO = "M"
    ElseIf Me.cboTipoProd.ListIndex = 2 Then 'Combo
        vTIPO = "C"
    ElseIf Me.cboTipoProd.ListIndex = 3 Then 'Materia Prima
        vTIPO = "I"
    End If

    CargarFamilias
    CargarDatos
End Sub

Private Sub DatFamilia_Change()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPRODUCTOS_LISTARsubFAMILIAS"

    Dim ORSsf As ADODB.Recordset

    Set ORSsf = oCmdEjec.Execute(, Array(LK_CODCIA, Me.DatFamilia.BoundText))

    Set Me.datSubFamilia.RowSource = ORSsf
    Me.datSubFamilia.BoundColumn = ORSsf.Fields(0).Name
    Me.datSubFamilia.ListField = ORSsf.Fields(1).Name
    Me.datSubFamilia.BoundText = "-1"
    CargarDatos
End Sub

Private Sub datSubFamilia_Change()
CargarDatos
End Sub

Private Sub Form_Load()
    ConfigurarLV

    For fila = 1 To lk_OTROS_Count

        If val(lk_OTROS(fila)) = 1 Then ' bloque de precios en mastros de articulos
            loc_flag_bloq = "A"
        End If

    Next fila

    Me.tbBotones.Buttons(2).Enabled = False
    Me.tbBotones.Buttons(3).Enabled = False

    If loc_flag_bloq = "A" Then
        Me.cboTipoProd.Locked = True
        Me.cboTipoProd.ListIndex = 2
    Else
        Me.cboTipoProd.ListIndex = 0
    End If

    VNuevo = False
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub CargarFormProd(EsNuevo As Boolean)
    frmProductos.VNuevo = EsNuevo
    'frmProductos.lblTipo.Caption = UCase(Me.cboTipoProd.Text)
    frmProductos.ComCboTipoProd.Text = UCase(Me.cboTipoProd.Text)

    If EsNuevo Then

        Select Case Me.cboTipoProd.ListIndex

            Case 0: frmProductos.cTipo = "P"

            Case 1: frmProductos.cTipo = "M"

            Case 2: frmProductos.cTipo = "C"

            Case 3: frmProductos.cTipo = "I"
        End Select

        frmProductos.Show vbModal
        CargarDatos
    Else

        If Not Me.lvData.SelectedItem Is Nothing Then
            'Dim oFrmM As New
            'frmProductos.vCodigo = Me.lvData.SelectedItem.Text
            frmProductos.txtDescripcion.Text = Me.lvData.SelectedItem.SubItems(1)
            '  frmProductos.txtUnidad.Text = Me.lvData.SelectedItem.SubItems(2)
            frmProductos.dcboFam.BoundText = Me.lvData.SelectedItem.SubItems(11)
            frmProductos.dcboSubFam.BoundText = Me.lvData.SelectedItem.SubItems(12)
            frmProductos.lblUniMedF.Caption = Me.lvData.SelectedItem.SubItems(2)
        
            Select Case Me.cboTipoProd.ListIndex

                Case 0: frmProductos.cTipo = "P"

                Case 1: frmProductos.cTipo = "M"

                Case 2: frmProductos.cTipo = "C"

                Case 3: frmProductos.cTipo = "I"
            End Select

            frmProductos.Show vbModal
            'CargarDatos
        Else
            MsgBox "Debe seleccionar un Articulo", vbInformation + vbOKOnly, "Error"
        End If
    End If

    If frmProductos.vGraba Then
        If frmProductos.vSit Then
            Me.lvData.SelectedItem.ForeColor = vbBlack
            Me.lvData.SelectedItem.Checked = True
            Me.lvData.SelectedItem.ListSubItems(1).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(2).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(3).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(4).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(5).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(6).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(7).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(8).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(9).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(10).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(11).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(12).ForeColor = vbBlack
            'Me.lvData.SelectedItem.ListSubItems(13).ForeColor = vbBlack
            'Me.lvData.SelectedItem.ListSubItems(14).ForeColor = vbBlack
        Else
            Me.lvData.SelectedItem.ForeColor = vbRed
            Me.lvData.SelectedItem.Checked = False
            Me.lvData.SelectedItem.ListSubItems(1).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(2).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(3).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(4).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(5).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(6).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(7).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(8).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(9).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(10).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(11).ForeColor = vbRed
            Me.lvData.SelectedItem.ListSubItems(12).ForeColor = vbRed
           Me.lvData.SelectedItem.ListSubItems(13).ForeColor = vbBlack
            Me.lvData.SelectedItem.ListSubItems(14).ForeColor = vbBlack
        End If
    End If

End Sub



Private Sub lvData_Click()
Me.tbBotones.Buttons(3).Enabled = True
End Sub

Private Sub lvData_DblClick()
frmProductos.vCodigo = Me.lvData.SelectedItem.Text
    CargarFormProd False

End Sub

Private Sub lvData_ItemCheck(ByVal Item As MSComctlLib.ListItem)

On Error GoTo Exito
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPPRODUCTOESTADO"
 oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adChar, adParamInput, 1, IIf(Item.Checked, "0", "1"))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adDouble, adParamInput, , Item)
    
    oCmdEjec.Execute
    
    If Item.Checked Then
    Me.lvData.ListItems(Item.index).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(1).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(2).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(3).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(4).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(5).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(6).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(7).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(8).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(9).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(10).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(11).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(12).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(13).ForeColor = vbBlack
    Me.lvData.ListItems(Item.index).ListSubItems(14).ForeColor = vbBlack
 
        Else
           Me.lvData.ListItems(Item.index).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(1).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(2).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(3).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(4).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(5).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(6).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(7).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(8).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(9).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(10).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(11).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(12).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(13).ForeColor = vbRed
    Me.lvData.ListItems(Item.index).ListSubItems(14).ForeColor = vbRed
        End If
    MsgBox "Datos Modificados correctamente."
    
     
        
    Exit Sub
Exito:
MsgBox Err.Description
End Sub

Private Sub tbBotones_ButtonClick(ByVal Button As ComctlLib.Button)

    Select Case Button.index

        Case 1
            CargarFormProd True
            frmProductos.VNuevo = True
            'frmProductos.lblTipo.Caption = UCase(Me.cboTipoProd.Text)
    
        Case 2
            CargarFormProd False

        Case 3 'ELIMINAR

            If MsgBox("¿Esta seguro que desea eliminar el ITEM seleccionado?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
        
            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SP_PRODUCTO_ELIMINA_VALIDA"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adBigInt, adParamInput, , Me.lvData.SelectedItem.Text)
        
            Dim ORSv As ADODB.Recordset

            Set ORSv = oCmdEjec.Execute
        
            If Not ORSv.EOF Then
                If CBool(ORSv!valida) Then
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SP_PRODUCTO_ELIMINA"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adBigInt, adParamInput, , Me.lvData.SelectedItem.Text)
                    oCmdEjec.Execute
                    Me.lvData.ListItems.Remove Me.lvData.SelectedItem.index
                    CargarDatos
                Else
                    MsgBox "No se puede eliminar el ITEM seleccionado.", vbCritical, Pub_Titulo
                End If
            End If

    End Select

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Len(Trim(Me.txtCodigo.Text)) = 0 Then
        oRsProd.Filter = ""
    Else
        oRsProd.Filter = "Codigo=" & Trim(Me.txtCodigo.Text)
    End If
    If Not oRsProd.EOF Then
        Me.lvData.ListItems.Clear
        Dim itemX As Object
        Do While Not oRsProd.EOF
            Set itemX = Me.lvData.ListItems.Add(, , oRsProd!Codigo)
            itemX.SubItems(1) = Trim(oRsProd!producto)
            itemX.SubItems(2) = Trim(oRsProd!UNIDAD)
            itemX.SubItems(3) = Trim(oRsProd!Familia)
            itemX.SubItems(4) = oRsProd.Fields(4)
            itemX.SubItems(5) = oRsProd.Fields(5)
            itemX.SubItems(6) = oRsProd.Fields(6)
            itemX.SubItems(7) = oRsProd.Fields(7)
            itemX.SubItems(8) = oRsProd.Fields(8)
            itemX.SubItems(9) = oRsProd.Fields(9)
            itemX.SubItems(10) = oRsProd!stock
            itemX.SubItems(11) = oRsProd!idfam
            itemX.SubItems(12) = oRsProd!idsubfam
            oRsProd.MoveNext
        Loop
    End If
End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        If Len(Trim(Me.txtDescripcion.Text)) = 0 Then
            oRsProd.Filter = ""
        Else
            oRsProd.Filter = "Producto like '%" & Me.txtDescripcion.Text & "%'"
        End If

        If Not oRsProd.EOF Then
            Me.lvData.ListItems.Clear

            Dim itemX As Object

            Do While Not oRsProd.EOF
                Set itemX = Me.lvData.ListItems.Add(, , oRsProd!Codigo)
                itemX.SubItems(1) = Trim(oRsProd!producto)
                itemX.SubItems(2) = Trim(oRsProd!UNIDAD)
                itemX.SubItems(3) = Trim(oRsProd!Familia)
                itemX.SubItems(4) = oRsProd.Fields(4)
                itemX.SubItems(5) = oRsProd.Fields(5)
                itemX.SubItems(6) = oRsProd.Fields(6)
                itemX.SubItems(7) = oRsProd.Fields(7)
                itemX.SubItems(8) = oRsProd.Fields(8)
                itemX.SubItems(9) = oRsProd.Fields(9)
                itemX.SubItems(10) = oRsProd!stock
                itemX.SubItems(11) = oRsProd!idfam
                itemX.SubItems(12) = oRsProd!idsubfam
                itemX.SubItems(13) = oRsProd!descontar
                itemX.SubItems(14) = oRsProd!formulacion

                If oRsProd!SIT = 1 Then
                    itemX.ForeColor = vbRed
                    itemX.ListSubItems(1).ForeColor = vbRed
                    itemX.ListSubItems(2).ForeColor = vbRed
                    itemX.ListSubItems(3).ForeColor = vbRed
                    itemX.ListSubItems(4).ForeColor = vbRed
                    itemX.ListSubItems(5).ForeColor = vbRed
                    itemX.ListSubItems(6).ForeColor = vbRed
                    itemX.ListSubItems(7).ForeColor = vbRed
                    itemX.ListSubItems(8).ForeColor = vbRed
                    itemX.ListSubItems(9).ForeColor = vbRed
                    itemX.ListSubItems(10).ForeColor = vbRed
                    itemX.ListSubItems(11).ForeColor = vbRed
                    itemX.ListSubItems(12).ForeColor = vbRed
                    itemX.ListSubItems(13).ForeColor = vbRed
                    itemX.ListSubItems(14).ForeColor = vbRed
                End If

                oRsProd.MoveNext
            Loop
       Me.lvData.SelectedItem.Selected = False
        End If

    End If

End Sub
