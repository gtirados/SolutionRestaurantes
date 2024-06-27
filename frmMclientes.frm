VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMclientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9915
   Begin TabDlg.SSTab stabMesa 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Clientes"
      TabPicture(0)   =   "frmMclientes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label10"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "ComActivo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDni"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNombres"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtApellidos"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboSexo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dtpFN"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dcCategoria"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDireccion"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtEmail"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtObs"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMclientes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvListado"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtSearch"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtObs 
         Height          =   285
         Left            =   2760
         TabIndex        =   24
         Tag             =   "X"
         Top             =   3720
         Width           =   6615
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         Tag             =   "X"
         Top             =   3360
         Width           =   6615
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   2760
         TabIndex        =   22
         Tag             =   "X"
         Top             =   3000
         Width           =   6615
      End
      Begin MSDataListLib.DataCombo dcCategoria 
         Height          =   315
         Left            =   2760
         TabIndex        =   20
         Top             =   2640
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpFN 
         Height          =   300
         Left            =   2760
         TabIndex        =   19
         Top             =   2280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   529
         _Version        =   393216
         Format          =   92864513
         CurrentDate     =   42317
      End
      Begin VB.ComboBox cboSexo 
         Height          =   315
         ItemData        =   "frmMclientes.frx":0038
         Left            =   2760
         List            =   "frmMclientes.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtApellidos 
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Tag             =   "X"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtNombres 
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Tag             =   "X"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtDni 
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Tag             =   "X"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmMclientes.frx":005B
         Left            =   2760
         List            =   "frmMclientes.frx":0065
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -73920
         TabIndex        =   1
         Top             =   480
         Width           =   8415
      End
      Begin MSComctlLib.ListView lvListado 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   3
         Top             =   840
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   6588
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria:"
         Height          =   195
         Left            =   1800
         TabIndex        =   21
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sexo:"
         Height          =   195
         Left            =   2160
         TabIndex        =   17
         Top             =   2040
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   1320
         TabIndex        =   16
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   3480
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   1800
         TabIndex        =   14
         Top             =   3120
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Nacimiento:"
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos:"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dni:"
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   2040
         TabIndex        =   6
         Top             =   4080
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombres:"
         Height          =   195
         Left            =   1800
         TabIndex        =   5
         Top             =   1245
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   4
         Top             =   480
         Width           =   675
      End
   End
   Begin MSComctlLib.Toolbar tbMesa 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   635
      ButtonWidth     =   2143
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ilMesa"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Guardar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "M&odificar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Desactiva"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Activa"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilMesa 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMclientes.frx":0071
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMclientes.frx":060B
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMclientes.frx":0BA5
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMclientes.frx":113F
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMclientes.frx":16D9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMclientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Sub Mandar_Datos()
With Me.lvListado
    Me.txtDni.Text = .SelectedItem.Text
    Me.txtNombres.Text = .SelectedItem.SubItems(1)
    Me.txtApellidos.Text = .SelectedItem.SubItems(2)
    Me.cboSexo.ListIndex = IIf(.SelectedItem.SubItems(3) = "M", 1, 0)
    Me.dtpFN.Value = .SelectedItem.SubItems(4)
    Me.txtDireccion.Text = .SelectedItem.SubItems(5)
    Me.dcCategoria.BoundText = .SelectedItem.SubItems(6)
    Me.txtEmail.Text = .SelectedItem.SubItems(7)
    Me.txtObs.Text = .SelectedItem.SubItems(8)
    Me.ComActivo.ListIndex = IIf(.SelectedItem.SubItems(9), 1, 0)
If .SelectedItem.SubItems(9) Then
Me.tbMesa.Buttons(5).Enabled = True
Me.tbMesa.Buttons(6).Enabled = False
Else
Me.tbMesa.Buttons(5).Enabled = False
Me.tbMesa.Buttons(6).Enabled = True
End If
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub Estado_Botones(val As Valores)

Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbMesa.Buttons(1).Enabled = True
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = False
        Me.tbMesa.Buttons(5).Enabled = False
        Me.tbMesa.Buttons(6).Enabled = False
        Me.stabMesa.tab = 0
    Case Nuevo, Editar
        Me.tbMesa.Buttons(1).Enabled = False
        Me.tbMesa.Buttons(2).Enabled = True
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = True
         Me.tbMesa.Buttons(5).Enabled = False
         Me.tbMesa.Buttons(6).Enabled = False
        Me.lvListado.Enabled = False
        Me.txtSearch.Enabled = False
        Me.stabMesa.tab = 0
    Case buscar
        Me.tbMesa.Buttons(1).Enabled = True
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = False
        Me.stabMesa.tab = 1
    Case AntesDeActualizar
        Me.tbMesa.Buttons(1).Enabled = False
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = True
        Me.tbMesa.Buttons(4).Enabled = True
'         Me.tbMesa.Buttons(5).Enabled = True
'         Me.tbMesa.Buttons(6).Enabled = True
        Me.stabMesa.tab = 0
End Select
End Sub

Private Sub Form_Load()
ConfigurarLV
Estado_Botones InicializarFormulario
DesactivarControles Me
ListarClientes
End Sub

Private Sub ConfigurarLV()
With Me.lvListado
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Dni"
    .ColumnHeaders.Add , , "Nombres", 4000
    .ColumnHeaders.Add , , "Apellidos", 300
    .ColumnHeaders.Add , , "sexo", 0
    .ColumnHeaders.Add , , "fechanacimiento", 0
    .ColumnHeaders.Add , , "direccion", 0
    .ColumnHeaders.Add , , "idcategoria", 0
    .ColumnHeaders.Add , , "email", 0
    .ColumnHeaders.Add , , "observaciones", 0
    .ColumnHeaders.Add , , "activo", 0
End With
End Sub


Private Sub ListarClientes()
Dim oRsZonas As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_CLIENTE_LISTAR"
Set oRsZonas = oCmdEjec.Execute(, Me.txtSearch.Text)

Dim tt As Object
Dim itemX As MSComctlLib.ListItem

Me.lvListado.ListItems.Clear
Do While Not oRsZonas.EOF
    Set itemX = Me.lvListado.ListItems.Add(, , oRsZonas!dni)
    itemX.SubItems(1) = oRsZonas!nombres
    itemX.SubItems(2) = oRsZonas!apellidos
    itemX.SubItems(3) = oRsZonas!sexo
    itemX.SubItems(4) = oRsZonas!fechanacimiento
    itemX.SubItems(5) = oRsZonas!direccion
    itemX.SubItems(6) = oRsZonas!idcategoria
    itemX.SubItems(7) = oRsZonas!email
    itemX.SubItems(8) = oRsZonas!observaciones
    itemX.SubItems(9) = oRsZonas!activo
    oRsZonas.MoveNext
Loop


End Sub





Private Sub lvListado_DblClick()
    Mandar_Datos
End Sub

Private Sub tbMesa_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            Me.ComActivo.ListIndex = 1
            
            Me.ComActivo.Enabled = False
            Me.txtDni.SetFocus

        Case 2 'Guardar
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtDni.Text)) = 0 Then
                MsgBox "Debe ingresar el DNI", vbCritical, NombreProyecto
                Me.txtDni.SetFocus

            Else

                On Error GoTo grabar

                Dim vCodigo As Integer

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRES", adVarChar, adParamInput, 80, Trim(Me.txtNombres.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@APELLIDOS", adVarChar, adParamInput, 80, Trim(Me.txtApellidos.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEXO", adChar, adParamInput, 1, IIf(Me.cboSexo.ListIndex = 0, "M", "H"))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHANAC", adDBTimeStamp, adParamInput, , Me.dtpFN.Value)
                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 200, Me.txtDireccion.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCATEGORIA", adBigInt, adParamInput, , -1) 'Me.dcCategoria.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EMAIL", adVarChar, adParamInput, 200, Me.txtEmail.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@OBS", adVarChar, adParamInput, 200, Me.txtObs.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CIAREGISTRO", adChar, adParamInput, 2, xParametros.CiaRegistro)

                If VNuevo Then
                    oCmdEjec.CommandText = "SP_CLIENTE_PUNTO_REGISTRAR"
                    
                Else

'                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMOTIVO", adInteger, adParamInput, , Me.lblID.Caption)
'                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Activo", adInteger, adParamInput, , Me.ComActivo.ListIndex)

                    oCmdEjec.CommandText = "SP_CLIENTE_PUNTO_MODIFICAR"
                End If

                oCmdEjec.Execute


                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
                'ListarZonas Me.txtBusMesa.Text
                Me.lvListado.Enabled = True
                Me.txtSearch.Enabled = True

'               ListarZonas Me.txtBusMesa.Text

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

            Me.ComActivo.Enabled = True
            Me.txtNombres.SetFocus

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvListado.Enabled = True
            Me.txtSearch.Enabled = True

        Case 5 'Desactiva

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then

                On Error GoTo elimina

                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "SP_CLIENTE_ESTADO"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , False)
                'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , CInt(Me.dcboZona.BoundText))
                oCmdEjec.Execute
                LimpiarControles Me
                Me.lvListado.Enabled = True
                'Me.lvListado.ListItems.Remove Me.lvListado.SelectedItem.Index
                Me.txtSearch.Enabled = True
                
                Estado_Botones Eliminar

                Exit Sub

elimina:
                MsgBox Err.Description, vbInformation, NombreProyecto
            End If
            Case 6 'Activa
             If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then

                On Error GoTo activa

                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "SP_CLIENTE_ESTADO"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Me.txtDni.Text)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , True)
                'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , CInt(Me.dcboZona.BoundText))
                oCmdEjec.Execute
                LimpiarControles Me
                Me.lvListado.Enabled = True
                'Me.lvListado.ListItems.Remove Me.lvListado.SelectedItem.Index
                Me.txtSearch.Enabled = True
                
                Estado_Botones Eliminar

                Exit Sub

activa:
                MsgBox Err.Description, vbInformation, NombreProyecto
            End If

    End Select

End Sub

