VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmContrata 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Contratas"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
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
   ScaleHeight     =   4605
   ScaleWidth      =   9285
   Begin TabDlg.SSTab sstbContrata 
      Height          =   3975
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Contrata"
      TabPicture(0)   =   "frmContrata.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCodigo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtRazonSocial"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtContacto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtFono"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ComActivo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmContrata.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBusMesa"
      Tab(1).Control(1)=   "lvContrata"
      Tab(1).Control(2)=   "Label4"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtBusMesa 
         Height          =   285
         Left            =   -73920
         TabIndex        =   13
         Top             =   360
         Width           =   7815
      End
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmContrata.frx":0038
         Left            =   4080
         List            =   "frmContrata.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtFono 
         Height          =   285
         Left            =   4080
         TabIndex        =   5
         Tag             =   "X"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtContacto 
         Height          =   285
         Left            =   4080
         TabIndex        =   4
         Tag             =   "X"
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtRazonSocial 
         Height          =   285
         Left            =   4080
         TabIndex        =   3
         Tag             =   "X"
         Top             =   1440
         Width           =   2655
      End
      Begin MSComctlLib.ListView lvContrata 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   12
         Top             =   720
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5530
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contrata:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   14
         Top             =   405
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   3315
         TabIndex        =   11
         Top             =   2700
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         Height          =   195
         Left            =   3105
         TabIndex        =   10
         Top             =   2205
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
         Height          =   195
         Left            =   3075
         TabIndex        =   9
         Top             =   1845
         Width           =   840
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social:"
         Height          =   195
         Left            =   2745
         TabIndex        =   8
         Top             =   1485
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   3240
         TabIndex        =   7
         Top             =   990
         Width           =   675
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Tag             =   "X"
         Top             =   960
         Width           =   1335
      End
   End
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
            Picture         =   "frmContrata.frx":004E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrata.frx":05E8
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrata.frx":0B82
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrata.frx":111C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmContrata.frx":16B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMesa 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9285
      _ExtentX        =   16378
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
            Object.Visible         =   0   'False
            Caption         =   "&Eliminar"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmContrata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean


Private Sub ListarContrata()
Dim oRScontrata As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_CONTRATA_LIST"

'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)

Set oRScontrata = oCmdEjec.Execute

Dim tt As Object
Dim itemX As MSComctlLib.ListItem

Me.lvContrata.ListItems.Clear
Do While Not oRScontrata.EOF
    Set itemX = Me.lvContrata.ListItems.Add(, , oRScontrata!Codigo)
    itemX.SubItems(1) = oRScontrata!RAZONSOCIAL
    itemX.SubItems(2) = oRScontrata!CONTACTO
    itemX.SubItems(3) = oRScontrata!activo
    
    oRScontrata.MoveNext
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
        Me.sstbContrata.tab = 0
    Case Nuevo, Editar
        Me.tbMesa.Buttons(1).Enabled = False
        Me.tbMesa.Buttons(2).Enabled = True
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = True
         Me.tbMesa.Buttons(5).Enabled = False
        Me.lvContrata.Enabled = False
        Me.txtBusMesa.Enabled = False
        Me.sstbContrata.tab = 0
    Case buscar
        Me.tbMesa.Buttons(1).Enabled = True
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = False
        Me.tbMesa.Buttons(4).Enabled = False
        Me.sstbContrata.tab = 1
    Case AntesDeActualizar
        Me.tbMesa.Buttons(1).Enabled = False
        Me.tbMesa.Buttons(2).Enabled = False
        Me.tbMesa.Buttons(3).Enabled = True
        Me.tbMesa.Buttons(4).Enabled = True
         Me.tbMesa.Buttons(5).Enabled = True
        Me.sstbContrata.tab = 0
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
ConfigurarLV
Estado_Botones InicializarFormulario
DesactivarControles Me
ListarContrata
End Sub



Private Sub lvContrata_DblClick()
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
            Me.txtRazonSocial.SetFocus

        Case 2 'Guardar
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtRazonSocial.Text)) = 0 Then
                MsgBox "Debe ingresar la Razón Social de la Contrata", vbCritical, Pub_Titulo
                Me.txtRazonSocial.SetFocus

            Else

                On Error GoTo grabar

                Dim vCodigo As Integer

                oCmdEjec.Prepared = True

                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@RAZONSOCIAL", adVarChar, adParamInput, 200, Trim(Me.txtRazonSocial.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CONTACTO", adVarChar, adParamInput, 200, Trim(Me.txtContacto.Text))

                If VNuevo Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATA", adInteger, adParamOutput, , vCodigo)
                    oCmdEjec.CommandText = "SP_CONTRATA_REGISTRAR"
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATA", adInteger, adParamInput, , Me.lblCodigo.Caption)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVO", adInteger, adParamInput, , Me.ComActivo.ListIndex)
                    oCmdEjec.CommandText = "SP_CONTRATA_MODIFICAR"
                End If

                oCmdEjec.Execute
                vCodigo = oCmdEjec.Parameters("@IDCONTRATA").Value

                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
                ListarContrata
                Me.lvContrata.Enabled = True
                Me.txtBusMesa.Enabled = True

                '        If VNuevo Then
                '            Me.txtCodigo.Text = vCodigo
                '            With Me.lvZonas.ListItems.Add(, , Me.txtCodigo.Text)
                '                .SubItems(1) = Trim(Me.txtDenominacion.Text)
                '            End With
                '        Else
                '            Me.lvZonas.SelectedItem.Text = Me.txtCodigo.Text
                '            Me.lvZonas.SelectedItem.SubItems(1) = Trim(Me.txtDenominacion.Text)
                '        End If

                'set itemg=me.lvMesas.ListItems.Add(,,
                MsgBox "Datos Almacenados Correctamente", vbInformation, Pub_Titulo

                Exit Sub

grabar:
                MsgBox Err.Description, vbInformation, Pub_Titulo

            End If

        Case 3 'Modificar
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
            
            Me.ComActivo.Enabled = True
            Me.txtRazonSocial.SetFocus

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvContrata.Enabled = True
            Me.txtBusMesa.Enabled = True

        Case 5 'Eliminar

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
''                Me.lvZonas.Enabled = True
''                Me.lvZonas.ListItems.Remove Me.lvZonas.SelectedItem.Index
''                Me.txtBusMesa.Enabled = True
''                Estado_Botones Eliminar
''
''                Exit Sub
''
''elimina:
''                MsgBox Err.Description, vbInformation, NombreProyecto
''            End If

    End Select
End Sub

Private Sub ConfigurarLV()
With Me.lvContrata
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CODIGO"
    .ColumnHeaders.Add , , "RAZONSOCIAL", 4000
    .ColumnHeaders.Add , , "CONTACTO", 1000
    .ColumnHeaders.Add , , "ACTIVO", 500
End With
End Sub


Sub Mandar_Datos()
With Me.lvContrata
    Me.lblCodigo.Caption = .SelectedItem.Text
    Me.txtRazonSocial.Text = .SelectedItem.SubItems(1)
    Me.txtContacto.Text = .SelectedItem.SubItems(2)
   If .SelectedItem.SubItems(3) = "SI" Then
    Me.ComActivo.ListIndex = 1
   Else
   Me.ComActivo.ListIndex = 2
   End If
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub txtBusMesa_Change()
Dim CantidadLetras As Integer
    CantidadLetras = Len(Trim(Me.txtBusMesa.Text))
    For i = 1 To Me.lvContrata.ListItems.count
        If Left(Me.lvContrata.ListItems(i).Text, CantidadLetras) = Trim(Me.txtBusMesa.Text) Then
            Me.lvContrata.ListItems(i).Selected = True
            Me.lvContrata.ListItems(i).EnsureVisible
            Exit For
        End If
    Next
End Sub

Private Sub txtBusMesa_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
If KeyAscii = vbKeyReturn Then Mandar_Datos
End Sub
