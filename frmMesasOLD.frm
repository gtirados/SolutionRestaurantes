VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMesas 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Mesas"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMesas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilMesa 
      Left            =   7680
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesas.frx":08CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesas.frx":0E64
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesas.frx":13FE
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMesas.frx":1998
            Key             =   "Undo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMesa 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ilMesa"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stabMesa 
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483636
      TabCaption(0)   =   "Mesa"
      TabPicture(0)   =   "frmMesas.frx":1F32
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDenominacion"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCodigo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dcboZona"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMesas.frx":1F4E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvMesas"
      Tab(1).Control(1)=   "txtBusMesa"
      Tab(1).Control(2)=   "Label5"
      Tab(1).ControlCount=   3
      Begin MSDataListLib.DataCombo dcboZona 
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   3000
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "X"
         Top             =   1200
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvMesas 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   9
         Top             =   840
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txtDenominacion 
         Height          =   285
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "X"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtBusMesa 
         Height          =   285
         Left            =   -74160
         TabIndex        =   8
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   2250
         TabIndex        =   10
         Top             =   1245
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mesa:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   7
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Zona:"
         Height          =   195
         Left            =   2415
         TabIndex        =   6
         Top             =   2205
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   1635
         TabIndex        =   5
         Top             =   1725
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmMesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Private Sub ListarZonas()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarZonas"
Dim oRsZ As ADODB.Recordset
Set oRsZ = oCmdEjec.Execute(, CodCia)
Set Me.dcboZona.RowSource = oRsZ
Me.dcboZona.ListField = oRsZ.Fields(1).Name
Me.dcboZona.BoundColumn = oRsZ.Fields(0).Name
End Sub

Sub Mandar_Datos()
With Me.lvMesas
Me.txtCodigo.Text = .SelectedItem.Tag
    Me.txtDenominacion.Text = .SelectedItem.Text
    'Me.txtDenominacion.Text = Trim(.SelectedItem.SubItems(1))
    'Me.txtZona.Text = Trim(.SelectedItem.SubItems(2))
    Me.dcboZona.BoundText = .SelectedItem.SubItems(1)
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub ConfigurarLv()
With Me.lvMesas
    .GridLines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Mesa"
    .ColumnHeaders.Add , , "CodZona", 0
    .ColumnHeaders.Add , , "Zona"
    .ColumnHeaders.Add , , "Estado", 0
End With
End Sub
Private Sub ListarMesas()
Dim oRsMesas As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarMesas"
Set oRsMesas = oCmdEjec.Execute(, CodCia)
Dim ItemM As ListItem
Do While Not oRsMesas.EOF
    Set ItemM = Me.lvMesas.ListItems.Add(, , Trim(oRsMesas!Mesa))
    ItemM.Tag = Trim(oRsMesas!codmesa)
    ItemM.SubItems(1) = oRsMesas!codZona
    ItemM.SubItems(2) = oRsMesas!Zona
    ItemM.SubItems(3) = Trim(oRsMesas!Estado)
    oRsMesas.MoveNext
Loop
End Sub

Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, Grabar, Cancelar
        Me.tbMesa.Buttons(1).Enabled = True
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = False
        Me.stabMesa.Tab = 0
    Case Nuevo, Editar
        Me.tbMesa.Buttons(1).Enabled = False
        Me.tbMesa.Buttons(2).Enabled = True
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = True
        Me.lvMesas.Enabled = False
        Me.txtBusMesa.Enabled = False
        Me.stabMesa.Tab = 0
    Case Buscar
        Me.tbMesa.Buttons(1).Enabled = True
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = False
        Me.stabMesa.Tab = 1
    Case AntesDeActualizar
        Me.tbMesa.Buttons(1).Enabled = False
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = True
        Me.tbMesa.Buttons(4).Enabled = True
        Me.stabMesa.Tab = 0
End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
ConfigurarLv
Estado_Botones InicializarFormulario
DesactivarControles Me
ListarMesas
ListarZonas
End Sub

Private Sub lvMesas_DblClick()
If Me.lvMesas.ListItems.Count <> 0 Then Mandar_Datos
End Sub

Private Sub tbMesa_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 'NUEVO
    ActivarControles Me
    LimpiarControles Me
    Estado_Botones Nuevo
    VNuevo = True
    Me.txtCodigo.SetFocus
Case 2 'Guardar
    LimpiaParametros oCmdEjec
    If Len(Trim(Me.txtCodigo.Text)) = 0 Then
        MsgBox "Debe ingresar el Código", vbCritical, NombreProyecto
        Me.txtCodigo.SetFocus
    ElseIf Len(Trim(Me.txtDenominacion.Text)) = 0 Then
        MsgBox "Debe ingresar la Denominación de la Mesa", vbCritical, NombreProyecto
        Me.txtDenominacion.SetFocus
    ElseIf Me.dcboZona.BoundText = "" Then
    
        MsgBox "Debe ingresar la zona de la Mesa.", vbCritical, NombreProyecto
        Me.dcboZona.SetFocus
    Else
        If VNuevo Then
            oCmdEjec.CommandText = "SpRegistrarMesa"
        Else
            oCmdEjec.CommandText = "SpModificarMesa"
        End If
        On Error GoTo Grabar
        
        oCmdEjec.Prepared = True
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, CodCia)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodMes", adVarChar, adParamInput, 10, Trim(Me.txtCodigo.Text))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Mesa", adVarChar, adParamInput, 40, Trim(Me.txtDenominacion.Text))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, 1, Trim(Me.dcboZona.BoundText))
        oCmdEjec.Execute
        'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
        DesactivarControles Me
        Estado_Botones Grabar
        Me.lvMesas.Enabled = True
        Me.txtBusMesa.Enabled = True
        Dim itemG As ListItem
        If VNuevo Then
            Set itemG = Me.lvMesas.ListItems.Add(, , Trim(Me.txtDenominacion.Text))
            itemG.Tag = Trim(Me.txtCodigo.Text)
            itemG.SubItems(1) = Me.dcboZona.BoundText
            itemG.SubItems(2) = Me.dcboZona.Text
        Else
            Me.lvMesas.SelectedItem.Text = Trim(Me.txtDenominacion.Text)
            Me.lvMesas.SelectedItem.SubItems(1) = Me.dcboZona.BoundText
            Me.lvMesas.SelectedItem.SubItems(2) = Me.dcboZona.Text
        End If
        
        'set itemg=me.lvMesas.ListItems.Add(,,
        MsgBox "Datos Almacenados Correctamente", vbInformation, NombreProyecto
        Exit Sub
Grabar:
    MsgBox Err.Description, vbInformation, NombreProyecto

    End If
Case 3 'Modificar
    VNuevo = False
    Estado_Botones Editar
    ActivarControles Me
    Me.txtCodigo.Enabled = False
    Me.txtDenominacion.SetFocus
Case 4 'Cancelar
    Estado_Botones Cancelar
    DesactivarControles Me
    Me.lvMesas.Enabled = True
    Me.txtBusMesa.Enabled = True
End Select
End Sub

Private Sub txtBusMesa_Change()
Dim CantidadLetras As Integer
    CantidadLetras = Len(Trim(Me.txtBusMesa.Text))
    For I = 1 To Me.lvMesas.ListItems.Count
        If Left(Me.lvMesas.ListItems(I).Text, CantidadLetras) = Trim(Me.txtBusMesa.Text) Then
            Me.lvMesas.ListItems(I).Selected = True
            Me.lvMesas.ListItems(I).EnsureVisible
            Exit For
        End If
    Next
End Sub

Private Sub txtBusMesa_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then Mandar_Datos
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then Me.txtDenominacion.SetFocus
End Sub

Private Sub txtDenominacion_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then Me.dcboZona.SetFocus
End Sub

Private Sub txtZona_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
End Sub
