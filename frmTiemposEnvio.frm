VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTiemposEnvio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tiempos de Envio"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
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
   ScaleHeight     =   4170
   ScaleWidth      =   7575
   Begin TabDlg.SSTab SSTTiempos 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483636
      TabCaption(0)   =   "Tiempos"
      TabPicture(0)   =   "frmTiemposEnvio.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(3)=   "lblCodigo"
      Tab(0).Control(4)=   "txtMinutos"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmTiemposEnvio.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lvTiempos"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtBusMesa"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtMinutos 
         Height          =   285
         Left            =   -72360
         MaxLength       =   10
         TabIndex        =   6
         Tag             =   "X"
         Top             =   1965
         Width           =   1215
      End
      Begin VB.TextBox txtBusMesa 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   5895
      End
      Begin MSComctlLib.ListView lvTiempos 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4683
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
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -72360
         TabIndex        =   9
         Tag             =   "X"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "minutos"
         Height          =   195
         Left            =   -71040
         TabIndex        =   8
         Top             =   2010
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   -73110
         TabIndex        =   5
         Top             =   1605
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo:"
         Height          =   195
         Left            =   500
         TabIndex        =   4
         Top             =   480
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo:"
         Height          =   195
         Left            =   -73200
         TabIndex        =   3
         Top             =   2010
         Width           =   705
      End
   End
   Begin MSComctlLib.ImageList iTiempos 
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
            Picture         =   "frmTiemposEnvio.frx":0038
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTiemposEnvio.frx":05D2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTiemposEnvio.frx":0B6C
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTiemposEnvio.frx":1106
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTiemposEnvio.frx":16A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbTiempos 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      ButtonWidth     =   2328
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "iTiempos"
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
End
Attribute VB_Name = "frmTiemposEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean
Private Sub Form_Load()
ConfigurarLV
Estado_Botones InicializarFormulario
DesactivarControles Me
'ListarMesas
ListarTiempos
End Sub

Private Sub ListarTiempos()
Dim oRsMesas As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_TIEMPOS_LISTAR"
Set oRsMesas = oCmdEjec.Execute(, LK_CODCIA)

Do While Not oRsMesas.EOF

    With Me.lvTiempos.ListItems.Add(, , Trim(oRsMesas!Codigo))
        
        .SubItems(1) = oRsMesas!DENOMINACION
        .Tag = oRsMesas!Min
    End With
   
    oRsMesas.MoveNext
Loop
End Sub

Private Sub ConfigurarLV()
With Me.lvTiempos
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Codigo"
    .ColumnHeaders.Add , , "Descripcion"
End With
End Sub

Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbTiempos.Buttons(1).Enabled = True
        Me.tbTiempos.Buttons(2).Enabled = False
        Me.tbTiempos.Buttons(3).Enabled = False
        Me.tbTiempos.Buttons(4).Enabled = False
        Me.tbTiempos.Buttons(5).Enabled = False
        Me.SSTTiempos.tab = 0
    Case Nuevo, Editar
        Me.tbTiempos.Buttons(1).Enabled = False
        Me.tbTiempos.Buttons(2).Enabled = True
        Me.tbTiempos.Buttons(3).Enabled = False
        Me.tbTiempos.Buttons(4).Enabled = True
         Me.tbTiempos.Buttons(5).Enabled = False
        Me.lvTiempos.Enabled = False
        Me.txtBusMesa.Enabled = False
        Me.SSTTiempos.tab = 0
    Case buscar
        Me.tbTiempos.Buttons(1).Enabled = True
        Me.tbTiempos.Buttons(2).Enabled = False
        Me.tbTiempos.Buttons(3).Enabled = False
        Me.tbTiempos.Buttons(4).Enabled = False
        Me.SSTTiempos.tab = 1
    Case AntesDeActualizar
        Me.tbTiempos.Buttons(1).Enabled = False
        Me.tbTiempos.Buttons(2).Enabled = False
        Me.tbTiempos.Buttons(3).Enabled = True
        Me.tbTiempos.Buttons(4).Enabled = True
         Me.tbTiempos.Buttons(5).Enabled = True
        Me.SSTTiempos.tab = 0
End Select
End Sub

Private Sub lvTiempos_DblClick()
With Me.lvTiempos
Me.lblCodigo.Caption = .SelectedItem.Text
    
    'Me.txtDenominacion.Text = Trim(.SelectedItem.SubItems(1))
    'Me.txtZona.Text = Trim(.SelectedItem.SubItems(2))
    Me.txtMinutos.Text = .SelectedItem.Tag
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub tbTiempos_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
    
            Estado_Botones Nuevo
            VNuevo = True
            Me.txtMinutos.SetFocus

        Case 2 'Guardar
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtMinutos.Text)) = 0 Then
                MsgBox "Debe ingresar el tiempo, vbCritical, NombreProyecto"
                Me.txtMinutos.SetFocus
    
            Else

                On Error GoTo grabar

                If VNuevo Then
                    oCmdEjec.CommandText = "SP_TIEMPOS_REGISTRAR"

                    Dim Xid As Integer

                    Xid = 0
                    oCmdEjec.Prepared = True
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MINUTOS", adDouble, adParamInput, , Trim(Me.txtMinutos.Text))
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CURRENTUSER", adVarChar, adParamInput, 20, LK_CODUSU)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTIEMPOENVIO", adInteger, adParamOutput, , Xid)
                    oCmdEjec.Execute
                    Xid = oCmdEjec.Parameters("@IDTIEMPOENVIO").Value
                Else
                    oCmdEjec.CommandText = "SP_TIEMPOS_MODIFICAR"
                    oCmdEjec.Prepared = True
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MINUTOS", adDouble, adParamInput, , Trim(Me.txtMinutos.Text))
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTIEMPOENVIO", adInteger, adParamInput, , Me.lblCodigo.Caption)
                    oCmdEjec.Execute
                End If
        
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
        
                Me.lvTiempos.Enabled = True
                Me.txtBusMesa.Enabled = True

                If VNuevo Then
        
                    With Me.lvTiempos.ListItems.Add(, , Xid)
            
                        .SubItems(1) = Trim(Me.txtMinutos.Text) & " MINUTOS"
                    End With
            
                Else
                    
                    Me.lvTiempos.SelectedItem.SubItems(1) = Trim(Me.txtMinutos.Text) & " MINUTOS"
                    Me.lvTiempos.SelectedItem.Tag = Trim(Me.txtMinutos.Text)
            
                End If
        
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
            Me.lblCodigo.Enabled = False
            Me.txtMinutos.SetFocus

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvTiempos.Enabled = True
            Me.txtBusMesa.Enabled = True
    
        Case 5 'Eliminar

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then

                On Error GoTo elimina
            
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "SP_TIEMPOS_ELIMINAR"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTIEMPOENVIO", adInteger, adParamInput, , Me.lblCodigo.Caption)
                oCmdEjec.Execute
                LimpiarControles Me
                
                Me.lvTiempos.Enabled = True
                Me.lvTiempos.ListItems.Remove Me.lvTiempos.SelectedItem.Index
                Me.txtBusMesa.Enabled = True
                Estado_Botones Eliminar

                Exit Sub

elimina:
                MsgBox Err.Description, vbInformation, Pub_Titulo
            End If

    End Select

End Sub

Private Sub txtBusMesa_Change()
Dim CantidadLetras As Integer
    CantidadLetras = Len(Trim(Me.txtBusMesa.Text))
    For i = 1 To Me.lvTiempos.ListItems.count
        If Left(Me.lvTiempos.ListItems(i).Text, CantidadLetras) = Trim(Me.txtBusMesa.Text) Then
            Me.lvTiempos.ListItems(i).Selected = True
            Me.lvTiempos.ListItems(i).EnsureVisible
            Exit For
        End If
    Next
End Sub
