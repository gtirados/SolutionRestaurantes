VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantMotivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Motivos de Anulación de Productos"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
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
   ScaleHeight     =   3690
   ScaleWidth      =   7470
   Begin TabDlg.SSTab stabMesa 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Motivo"
      TabPicture(0)   =   "frmMantMotivos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ComActivo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDenominacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmMantMotivos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtBusMesa"
      Tab(1).Control(1)=   "lvMotivos"
      Tab(1).Control(2)=   "Label5"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtBusMesa 
         Height          =   285
         Left            =   -74160
         TabIndex        =   7
         Top             =   480
         Width           =   6255
      End
      Begin VB.TextBox txtDenominacion 
         Height          =   285
         Left            =   2805
         MaxLength       =   40
         TabIndex        =   3
         Tag             =   "X"
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmMantMotivos.frx":0038
         Left            =   2805
         List            =   "frmMantMotivos.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvMotivos 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   6
         Top             =   840
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   3836
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
      Begin VB.Label lblID 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   315
         Left            =   2760
         TabIndex        =   9
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   8
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   1245
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   2085
         TabIndex        =   4
         Top             =   1680
         Width           =   600
      End
   End
   Begin MSComctlLib.Toolbar tbMesa 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   635
      ButtonWidth     =   2037
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
            Picture         =   "frmMantMotivos.frx":004E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantMotivos.frx":05E8
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantMotivos.frx":0B82
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantMotivos.frx":111C
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMantMotivos.frx":16B6
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMantMotivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Sub Mandar_Datos()
With Me.lvMotivos
    Me.lblID.Caption = .SelectedItem.Text
    Me.txtDenominacion.Text = .SelectedItem.SubItems(1)
    Me.ComActivo.ListIndex = .SelectedItem.SubItems(3)
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub ConfigurarLV()
With Me.lvMotivos
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

Private Sub ListarZonas(xSEARCH As String)

    Dim oRsZonas As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_MOTIVOS_LISTAR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)

    If Len(Trim(xSEARCH)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SEARCH", adVarChar, adParamInput, 100, xSEARCH)
    End If
                
    Set oRsZonas = oCmdEjec.Execute

    Dim tt    As Object

    Dim itemX As MSComctlLib.ListItem

    Me.lvMotivos.ListItems.Clear

    Do While Not oRsZonas.EOF
        Set itemX = Me.lvMotivos.ListItems.Add(, , oRsZonas!cod)
        itemX.SubItems(1) = oRsZonas!MOTIVO
        itemX.SubItems(2) = oRsZonas!ACTIVO
        itemX.SubItems(3) = oRsZonas!ACT
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
        Me.lvMotivos.Enabled = False
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
ListarZonas Me.txtBusMesa.Text
End Sub

Private Sub lvMotivos_DblClick()
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
            Me.lblID.Caption = ""
            Me.ComActivo.Enabled = False
            Me.txtDenominacion.SetFocus

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
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCRIPCION", adVarChar, adParamInput, 40, Trim(Me.txtDenominacion.Text))
        
                If VNuevo Then
                    oCmdEjec.CommandText = "SP_MOTIVOS_REGISTRAR"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMOTIVO", adInteger, adParamOutput, , 0)
                Else
                    
                    
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMOTIVO", adInteger, adParamInput, , Me.lblID.Caption)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Activo", adInteger, adParamInput, , Me.ComActivo.ListIndex)
                    
                    oCmdEjec.CommandText = "SP_MOTIVOS_MODIFICAR"
                End If
        
                oCmdEjec.Execute
                
        
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
                ListarZonas Me.txtBusMesa.Text
                Me.lvMotivos.Enabled = True
                Me.txtBusMesa.Enabled = True

               ListarZonas Me.txtBusMesa.Text
        
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
            Me.txtDenominacion.SetFocus

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvMotivos.Enabled = True
            Me.txtBusMesa.Enabled = True

        Case 5 'Eliminar

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then

                On Error GoTo elimina
   
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "SP_MOTIVOS_ELIMINAR"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDMOTIVO", adInteger, adParamInput, , CInt(Me.lblID.Caption))
                'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , CInt(Me.dcboZona.BoundText))
                oCmdEjec.Execute
                LimpiarControles Me
                Me.lvMotivos.Enabled = True
                Me.lvMotivos.ListItems.Remove Me.lvMotivos.SelectedItem.Index
                Me.txtBusMesa.Enabled = True
                Me.lblID.Caption = ""
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

    For i = 1 To Me.lvMotivos.ListItems.count

        If Left(Me.lvMotivos.ListItems(i).SubItems(1), CantidadLetras) = Trim(Me.txtBusMesa.Text) Then
            Me.lvMotivos.ListItems(i).Selected = True
            Me.lvMotivos.ListItems(i).EnsureVisible

            Exit For

        End If

    Next

End Sub
