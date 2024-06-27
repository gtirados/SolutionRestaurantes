VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantZonas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Zonas"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7665
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMantZonas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   7665
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantZonas.frx":08CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantZonas.frx":0E64
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantZonas.frx":13FE
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantZonas.frx":1998
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantZonas.frx":1F32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMesa 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
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
            Caption         =   "&Eliminar"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stabMesa 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483636
      TabCaption(0)   =   "Zonas"
      TabPicture(0)   =   "frmMantZonas.frx":262D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "ComActivo"
      Tab(0).Control(1)=   "txtCodigo"
      Tab(0).Control(2)=   "txtDenominacion"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(5)=   "Label2"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMantZonas.frx":2649
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtBusMesa"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lvZonas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmMantZonas.frx":2665
         Left            =   -72000
         List            =   "frmMantZonas.frx":266F
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   -72000
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "X"
         Top             =   1200
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvZonas 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
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
         Left            =   -72000
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "X"
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox txtBusMesa 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   -72720
         TabIndex        =   9
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   -72720
         TabIndex        =   8
         Top             =   1245
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Mesa:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   -73365
         TabIndex        =   4
         Top             =   1725
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmMantZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Sub Mandar_Datos()
With Me.lvZonas
    Me.txtCodigo.Text = .SelectedItem.Text
    Me.txtDenominacion.Text = .SelectedItem.SubItems(1)
    Me.ComActivo.ListIndex = .SelectedItem.SubItems(3)
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub ConfigurarLV()
With Me.lvZonas
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Codigo"
    .ColumnHeaders.Add , , "Denominacion", 4000
    .ColumnHeaders.Add , , "Activo", 300
    .ColumnHeaders.Add , , "Valor", 0
End With
End Sub
Private Sub ListarZonas()
Dim oRsZonas As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SpListarZonas"
Set oRsZonas = oCmdEjec.Execute(, Array(LK_CODCIA, 1))

Dim tt As Object
Dim itemX As MSComctlLib.ListItem

Me.lvZonas.ListItems.Clear
Do While Not oRsZonas.EOF
    Set itemX = Me.lvZonas.ListItems.Add(, , oRsZonas!Codigo)
    itemX.SubItems(1) = oRsZonas!denomina
    itemX.SubItems(2) = oRsZonas!ACTIVO
    itemX.SubItems(3) = oRsZonas!valor
    oRsZonas.MoveNext
Loop


End Sub

Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbMesa.Buttons(1).Enabled = True
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = False
        Me.tbMesa.Buttons(5).Enabled = False
        Me.stabMesa.tab = 0
    Case Nuevo, Editar
        Me.tbMesa.Buttons(1).Enabled = False
        Me.tbMesa.Buttons(2).Enabled = True
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = True
         Me.tbMesa.Buttons(5).Enabled = False
        Me.lvZonas.Enabled = False
        Me.txtBusMesa.Enabled = False
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
         Me.tbMesa.Buttons(5).Enabled = True
        Me.stabMesa.tab = 0
End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
ConfigurarLV
Estado_Botones InicializarFormulario
DesactivarControles Me
ListarZonas
End Sub

Private Sub lvMesas_Click()
Mandar_Datos
End Sub

Private Sub lvMesas_DblClick()
If Me.lvZonas.ListItems.count <> 0 Then Mandar_Datos

End Sub

Private Sub lvZonas_Click()
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
    Me.txtCodigo.SetFocus
Case 2 'Guardar
    LimpiaParametros oCmdEjec
  If Len(Trim(Me.txtDenominacion.Text)) = 0 Then
        MsgBox "Debe ingresar la Denominación de la Mesa", vbCritical, NombreProyecto
        Me.txtDenominacion.SetFocus
   
    Else
    
        On Error GoTo grabar
        Dim vCodigo As Integer
        oCmdEjec.Prepared = True
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Descrip", adVarChar, adParamInput, 40, Trim(Me.txtDenominacion.Text))
        
            If VNuevo Then
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamOutput, , vCodigo)
            oCmdEjec.CommandText = "SpRegistrarZona"
        Else
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , Me.txtCodigo.Text)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Activo", adInteger, adParamInput, , Me.ComActivo.ListIndex)
            oCmdEjec.CommandText = "SpModificarZona"
        End If
        
        oCmdEjec.Execute
        vCodigo = oCmdEjec.Parameters("@Codigo").Value
        
        'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
        DesactivarControles Me
        Estado_Botones grabar
        ListarZonas
        Me.lvZonas.Enabled = True
        Me.txtBusMesa.Enabled = True

        If VNuevo Then
            Me.txtCodigo.Text = vCodigo
            With Me.lvZonas.ListItems.Add(, , Me.txtCodigo.Text)
                .SubItems(1) = Trim(Me.txtDenominacion.Text)
            End With
        Else
            Me.lvZonas.SelectedItem.Text = Me.txtCodigo.Text
            Me.lvZonas.SelectedItem.SubItems(1) = Trim(Me.txtDenominacion.Text)
        End If
        
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
    Me.txtDenominacion.SetFocus
Case 4 'Cancelar
    Estado_Botones cancelar
    DesactivarControles Me
    Me.lvZonas.Enabled = True
    Me.txtBusMesa.Enabled = True
Case 5 'Eliminar
If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
    On Error GoTo elimina
   
        LimpiaParametros oCmdEjec
         oCmdEjec.Prepared = True
        oCmdEjec.CommandText = "SpEliminarZona"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , CInt(Me.txtCodigo.Text))
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , CInt(Me.dcboZona.BoundText))
        oCmdEjec.Execute
        LimpiarControles Me
    Me.lvZonas.Enabled = True
    Me.lvZonas.ListItems.Remove Me.lvZonas.SelectedItem.Index
    Me.txtBusMesa.Enabled = True
    Estado_Botones Eliminar
    Exit Sub
elimina:
        MsgBox Err.Description, vbInformation, NombreProyecto
    End If
End Select
End Sub

Private Sub txtBusMesa_Change()
Dim CantidadLetras As Integer
    CantidadLetras = Len(Trim(Me.txtBusMesa.Text))
    For i = 1 To Me.lvZonas.ListItems.count
        If Left(Me.lvZonas.ListItems(i).Text, CantidadLetras) = Trim(Me.txtBusMesa.Text) Then
            Me.lvZonas.ListItems(i).Selected = True
            Me.lvZonas.ListItems(i).EnsureVisible
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
End Sub

Private Sub txtZona_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
End Sub
