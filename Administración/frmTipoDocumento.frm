VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form frmTipoDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Documento"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
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
   ScaleHeight     =   5115
   ScaleWidth      =   8970
   Begin TabDlg.SSTab stabTD 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Documentos"
      TabPicture(0)   =   "frmTipoDocumento.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDeno"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ComActivo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "comDefecto"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "comEditable"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtConsumo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDetallado"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmTipoDocumento.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtSearch"
      Tab(1).Control(1)=   "lvDatos"
      Tab(1).Control(2)=   "Label6"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtDetallado 
         Height          =   285
         Left            =   2520
         TabIndex        =   16
         Tag             =   "X"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox txtConsumo 
         Height          =   285
         Left            =   2520
         TabIndex        =   15
         Tag             =   "X"
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   -73560
         TabIndex        =   12
         Top             =   480
         Width           =   7215
      End
      Begin VB.ComboBox comEditable 
         Height          =   315
         ItemData        =   "frmTipoDocumento.frx":0038
         Left            =   2520
         List            =   "frmTipoDocumento.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox comDefecto 
         Height          =   315
         ItemData        =   "frmTipoDocumento.frx":004E
         Left            =   5400
         List            =   "frmTipoDocumento.frx":0058
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmTipoDocumento.frx":0064
         Left            =   2520
         List            =   "frmTipoDocumento.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtDeno 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Tag             =   "X"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Tag             =   "X"
         Top             =   960
         Width           =   1575
      End
      Begin MSComctlLib.ListView lvDatos 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   14
         Top             =   840
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6376
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detallado:"
         Height          =   195
         Left            =   1545
         TabIndex        =   18
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consumo:"
         Height          =   195
         Left            =   1545
         TabIndex        =   17
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   13
         Top             =   525
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Editable:"
         Height          =   195
         Left            =   1680
         TabIndex        =   11
         Top             =   2820
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defecto:"
         Height          =   195
         Left            =   4560
         TabIndex        =   5
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   1830
         TabIndex        =   4
         Top             =   2460
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   1680
         TabIndex        =   3
         Top             =   1320
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   1755
         TabIndex        =   2
         Top             =   960
         Width           =   675
      End
   End
   Begin MSComctlLib.Toolbar tbTD 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Sub Mandar_Datos()
With Me.lvDatos
    Me.txtCodigo.Text = .SelectedItem.Text
    Me.txtDeno.Text = .SelectedItem.SubItems(1)
    Me.ComActivo.ListIndex = IIf(.SelectedItem.SubItems(2) = "SI", 1, 0)
    Me.comDefecto.ListIndex = IIf(.SelectedItem.SubItems(3) = "SI", 1, 0)
    Me.comEditable.ListIndex = IIf(.SelectedItem.SubItems(4) = "SI", 1, 0)
Me.txtConsumo.Text = .SelectedItem.SubItems(5)
Me.txtDetallado.Text = .SelectedItem.SubItems(6)
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub ListarTD()
Dim oRsZonas As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_TIPOSDOCTOS_LISTAR"

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRE", adVarChar, adParamInput, 80, Me.txtSearch.Text)

Set oRsZonas = oCmdEjec.Execute

Dim tt As Object
Dim itemX As MSComctlLib.ListItem

Me.lvDatos.ListItems.Clear
Do While Not oRsZonas.EOF
    Set itemX = Me.lvDatos.ListItems.Add(, , oRsZonas!Codigo)
    itemX.SubItems(1) = oRsZonas!Nombre
    itemX.SubItems(2) = IIf(oRsZonas!ACTIVO, "SI", "NO")
    itemX.SubItems(3) = IIf(oRsZonas!defecto, "SI", "NO")
    itemX.SubItems(4) = IIf(oRsZonas!Editable, "SI", "NO")
    itemX.SubItems(5) = oRsZonas!CONSUMO
    itemX.SubItems(6) = oRsZonas!DETALLADO
    oRsZonas.MoveNext
Loop


End Sub

Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbTD.Buttons(1).Enabled = True
        Me.tbTD.Buttons(2).Enabled = False
        Me.tbTD.Buttons(3).Enabled = False
        Me.tbTD.Buttons(4).Enabled = False
       ' Me.tbTD.Buttons(5).Enabled = False
        Me.stabTD.tab = 0
    Case Nuevo, Editar
        Me.tbTD.Buttons(1).Enabled = False
        Me.tbTD.Buttons(2).Enabled = True
        Me.tbTD.Buttons(3).Enabled = False
        Me.tbTD.Buttons(4).Enabled = True
        ' Me.tbTD.Buttons(5).Enabled = False
        Me.lvDatos.Enabled = False
        Me.txtSearch.Enabled = False
        Me.stabTD.tab = 0
    Case buscar
        Me.tbTD.Buttons(1).Enabled = True
        Me.tbTD.Buttons(2).Enabled = False
        Me.tbTD.Buttons(3).Enabled = False
        Me.tbTD.Buttons(4).Enabled = False
        Me.stabTD.tab = 1
    Case AntesDeActualizar
        Me.tbTD.Buttons(1).Enabled = False
        Me.tbTD.Buttons(2).Enabled = False
        Me.tbTD.Buttons(3).Enabled = True
        Me.tbTD.Buttons(4).Enabled = True
        ' Me.tbTD.Buttons(5).Enabled = True
        Me.stabTD.tab = 0
End Select
End Sub


Private Sub Form_Load()
CenterMe Me
ConfigurarLV
Estado_Botones InicializarFormulario
DesactivarControles Me
ListarTD
End Sub
Private Sub ConfigurarLV()
With Me.lvDatos
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CODIGO"
    .ColumnHeaders.Add , , "DESCRIPCION", 4000
    .ColumnHeaders.Add , , "ACTIVO", 700
    .ColumnHeaders.Add , , "DEFECTO", 700
    .ColumnHeaders.Add , , "EDITABLE", 700
    .ColumnHeaders.Add , , "CONSUMO", 0
    .ColumnHeaders.Add , , "DETALLADO", 0
End With
End Sub

Private Sub lvDatos_DblClick()
Mandar_Datos
End Sub

Private Sub tbTD_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            'Me.txtCodigo.Enabled = False
            Estado_Botones Nuevo
            VNuevo = True
            Me.ComActivo.ListIndex = 1
            Me.txtCodigo.SetFocus

        Case 2 'Guardar
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtDeno.Text)) = 0 Then
                MsgBox "Debe ingresar la Denominación del turno", vbCritical, Pub_Titulo
                Me.txtDeno.SetFocus
'            ElseIf Len(Trim(Me.txtabreviatura.Text)) = 0 Then
'                MsgBox "Debe ingresar la Abreviatura del Turno", vbCritical, Pub_Titulo
'                Me.txtabreviatura.SetFocus
'            ElseIf Not IsNumeric(Me.txtTolerancia.Text) Then
'                MsgBox "La Tolerancia ingresada no es correcta.", vbCritical, Pub_Titulo
'                Me.txtTolerancia.SetFocus
'                Me.txtTolerancia.SelStart = 0
'                Me.txtTolerancia.SelLength = Len(Me.txtTolerancia.Text)
'            ElseIf val(Me.txtTolerancia.Text) > 60 Then
'                MsgBox "La Tolerancia supera los 60 minutos.", vbCritical, Pub_Titulo
'                Me.txtTolerancia.SetFocus
'                Me.txtTolerancia.SelStart = 0
'                Me.txtTolerancia.SelLength = Len(Me.txtTolerancia.Text)
'
'                Exit Sub
   
            Else
    
                On Error GoTo grabar

                Dim vCodigo As Integer

                oCmdEjec.Prepared = True
        
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adChar, adParamInput, 2, Me.txtCodigo.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRE", adVarChar, adParamInput, 80, Trim(Me.txtDeno.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVO", adBoolean, adParamInput, , Me.ComActivo.ListIndex)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DEFECTO", adBoolean, adParamInput, , Me.comDefecto.ListIndex)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EDITABLE", adBoolean, adParamInput, , Me.comEditable.ListIndex)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CONSUMO", adVarChar, adParamInput, 50, Me.txtConsumo.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DETALLADO", adVarChar, adParamInput, 50, Me.txtDetallado.Text)
        
                If VNuevo Then
                    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamOutput, , vCodigo)
                    oCmdEjec.CommandText = "SP_TIPOSDOCTOS_REGISTRAR"
                Else
                    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , Me.txtCodigo.Text)
                    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Activo", adInteger, adParamInput, , Me.ComActivo.ListIndex)
                    oCmdEjec.CommandText = "SP_TIPOSDOCTOS_MODIFICAR"
                End If
        
                oCmdEjec.Execute
               ' vCodigo = oCmdEjec.Parameters("@Codigo").Value
        
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
               ListarTD
                Me.lvDatos.Enabled = True
                Me.txtSearch.Enabled = True

         
        
                'set itemg=me.lvMesas.ListItems.Add(,,
                MsgBox "Datos Almacenados Correctamente", vbInformation, NombreProyecto

                Exit Sub

grabar:
                MsgBox Err.Description, vbInformation, NombreProyecto

            End If

        Case 3 'Modificar
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
            Me.txtCodigo.Enabled = False
            Me.ComActivo.Enabled = True
            Me.txtDeno.SetFocus

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvDatos.Enabled = True
            Me.txtSearch.Enabled = True

''        Case 5 'Eliminar
''
''            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
''
''                On Error GoTo elimina
''
''                LimpiaParametros oCmdEjec
''                oCmdEjec.Prepared = True
''                oCmdEjec.CommandText = "SpEliminarZona"
''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , CInt(Me.txtCodigo.Text))
''                'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , CInt(Me.dcboZona.BoundText))
''                oCmdEjec.Execute
''                LimpiarControles Me
''                Me.lvDatos.Enabled = True
''                Me.lvDatos.ListItems.Remove Me.lvDatos.SelectedItem.Index
''                Me.txtSearch.Enabled = True
''                Estado_Botones Eliminar
''
''                Exit Sub
''
''elimina:
''                MsgBox Err.Description, vbInformation, NombreProyecto
''            End If

    End Select
End Sub


Private Sub txtSearch_Change()
Dim CantidadLetras As Integer
    CantidadLetras = Len(Trim(Me.txtSearch.Text))
    For i = 1 To Me.lvDatos.ListItems.count
        If Left(Me.lvDatos.ListItems(i).SubItems(1), CantidadLetras) = Trim(Me.txtSearch.Text) Then
            Me.lvDatos.ListItems(i).Selected = True
            Me.lvDatos.ListItems(i).EnsureVisible
            Exit For
        End If
    Next
End Sub
