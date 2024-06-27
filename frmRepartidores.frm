VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRepartidores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Repartidores"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10305
   Begin TabDlg.SSTab SSTRepartidores 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8493
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
      TabCaption(0)   =   "Repartidor"
      TabPicture(0)   =   "frmRepartidores.frx":0000
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
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCodigo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtRepartidor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDni"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDireccion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtFono1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtFono2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPlaca"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtMarca"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtModelo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmRepartidores.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvRepartidores"
      Tab(1).Control(1)=   "txtSearch"
      Tab(1).Control(2)=   "Label10"
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView lvRepartidores 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   22
         Top             =   840
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   6800
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
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   -73680
         TabIndex        =   21
         Top             =   480
         Width           =   8655
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   8
         Tag             =   "X"
         Top             =   3480
         Width           =   5655
      End
      Begin VB.TextBox txtMarca 
         Height          =   285
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "X"
         Top             =   3120
         Width           =   5655
      End
      Begin VB.TextBox txtPlaca 
         Height          =   285
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "X"
         Top             =   2760
         Width           =   5655
      End
      Begin VB.TextBox txtFono2 
         Height          =   285
         Left            =   6840
         MaxLength       =   30
         TabIndex        =   5
         Tag             =   "X"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtFono1 
         Height          =   285
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   4
         Tag             =   "X"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Tag             =   "X"
         Top             =   2040
         Width           =   5655
      End
      Begin VB.TextBox txtDni 
         Height          =   285
         Left            =   3120
         MaxLength       =   8
         TabIndex        =   2
         Tag             =   "X"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtRepartidor 
         Height          =   285
         Left            =   3120
         MaxLength       =   150
         TabIndex        =   1
         Tag             =   "X"
         Top             =   1320
         Width           =   5655
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Repartidor:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   20
         Top             =   525
         Width           =   975
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Tag             =   "X"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Placa Vehiculo:"
         Height          =   195
         Left            =   1650
         TabIndex        =   18
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marca Vehiculo:"
         Height          =   195
         Left            =   1590
         TabIndex        =   17
         Top             =   3120
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modelo Vehiculo:"
         Height          =   195
         Left            =   1500
         TabIndex        =   16
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         Height          =   195
         Left            =   2085
         TabIndex        =   15
         Top             =   2040
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dni:"
         Height          =   195
         Left            =   2595
         TabIndex        =   14
         Top             =   1680
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono 1:"
         Height          =   195
         Left            =   1980
         TabIndex        =   13
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono 2:"
         Height          =   195
         Left            =   5640
         TabIndex        =   12
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellidos y Nombres:"
         Height          =   195
         Left            =   1125
         TabIndex        =   11
         Top             =   1320
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   2280
         TabIndex        =   10
         Top             =   960
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList iRepartidores 
      Left            =   9240
      Top             =   4920
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
            Picture         =   "frmRepartidores.frx":0038
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepartidores.frx":05D2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepartidores.frx":0B6C
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepartidores.frx":1106
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepartidores.frx":16A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbRepartidores 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   635
      ButtonWidth     =   2143
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "iRepartidores"
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
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Activa"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRepartidores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Private Sub ConfigurarLV()
With Me.lvRepartidores
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Repartidor", 3500
    .ColumnHeaders.Add , , "Dirección", 2500
    .ColumnHeaders.Add , , "Telf. 1"
    .ColumnHeaders.Add , , "Telf. 2"
    .ColumnHeaders.Add , , "Activo", 700
End With
End Sub

Private Sub ListarRepartidores()
Dim oRSr As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_REPARTIDORES_LISTAR"
Set oRSr = oCmdEjec.Execute(, LK_CODCIA)

Do While Not oRSr.EOF

    With Me.lvRepartidores.ListItems.Add(, , Trim(oRSr!REP))
        .Tag = Trim(oRSr!IDE)
        .SubItems(1) = oRSr!dir
        .SubItems(2) = oRSr!f1
        .SubItems(3) = oRSr!F2
        .SubItems(4) = oRSr!ac
    End With
   
    oRSr.MoveNext
Loop
End Sub

Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbRepartidores.Buttons(1).Enabled = True
        Me.tbRepartidores.Buttons(2).Enabled = False
        Me.tbRepartidores.Buttons(3).Enabled = False
        Me.tbRepartidores.Buttons(4).Enabled = False
        Me.tbRepartidores.Buttons(5).Enabled = False
        Me.tbRepartidores.Buttons(6).Enabled = False
        Me.SSTRepartidores.tab = 0
    Case Nuevo, Editar
        Me.tbRepartidores.Buttons(1).Enabled = False
        Me.tbRepartidores.Buttons(2).Enabled = True
        Me.tbRepartidores.Buttons(3).Enabled = False
        Me.tbRepartidores.Buttons(4).Enabled = True
        Me.tbRepartidores.Buttons(5).Enabled = False
         Me.tbRepartidores.Buttons(6).Enabled = False
        Me.lvRepartidores.Enabled = False
        Me.txtSearch.Enabled = False
        Me.SSTRepartidores.tab = 0
    Case buscar
        Me.tbRepartidores.Buttons(1).Enabled = True
        Me.tbRepartidores.Buttons(2).Enabled = False
        Me.tbRepartidores.Buttons(3).Enabled = False
        Me.tbRepartidores.Buttons(4).Enabled = False
        Me.SSTRepartidores.tab = 1
    Case AntesDeActualizar
        Me.tbRepartidores.Buttons(1).Enabled = False
        Me.tbRepartidores.Buttons(2).Enabled = False
        Me.tbRepartidores.Buttons(3).Enabled = True
        Me.tbRepartidores.Buttons(4).Enabled = True
         Me.tbRepartidores.Buttons(5).Enabled = True
          Me.tbRepartidores.Buttons(6).Enabled = True
        Me.SSTRepartidores.tab = 0
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
ListarRepartidores
'ListarZonas
End Sub

Sub Mandar_Datos()
With Me.lvRepartidores
Me.lblCodigo.Caption = .SelectedItem.Tag
    Me.txtRepartidor.Text = .SelectedItem.Text
    'Me.txtDenominacion.Text = Trim(.SelectedItem.SubItems(1))
    'Me.txtZona.Text = Trim(.SelectedItem.SubItems(2))
    Me.txtDireccion.Text = .SelectedItem.SubItems(1)
    Me.txtFono1.Text = .SelectedItem.SubItems(2)
    Me.txtFono2.Text = .SelectedItem.SubItems(3)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REPARTIDORES_FILL"
    oCmdEjec.CommandType = adCmdStoredProc
    Dim ORSd As ADODB.Recordset
    
    Set ORSd = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lblCodigo.Caption))
    
    If Not ORSd.EOF Then
        Me.txtDni.Text = ORSd!dni
        Me.txtMarca.Text = ORSd!marca
        Me.txtModelo.Text = ORSd!MODELO
        Me.txtPlaca.Text = ORSd!PLACA
    End If
    
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub lvRepartidores_DblClick()
Mandar_Datos
End Sub

Private Sub lvRepartidores_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then Mandar_Datos

End Sub

Private Sub SSTRepartidores_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
    txtSearch_KeyPress vbKeyReturn
End If
End Sub

Private Sub tbRepartidores_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            Me.txtRepartidor.SetFocus

        Case 2 'Guardar
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtRepartidor.Text)) = 0 Then
                MsgBox "Debe ingresar el Código", vbCritical, NombreProyecto
                Me.txtRepartidor.SetFocus
            ElseIf Len(Trim(Me.txtDni.Text)) = 0 Then
                MsgBox "Debe ingresar el DNI", vbCritical, NombreProyecto
                Me.txtDni.SetFocus
            ElseIf Len(Trim(Me.txtPlaca.Text)) = 0 Then
                MsgBox "Debe ingresar la Placa del Vehiculo"
                Me.txtPlaca.SetFocus
            Else

                If VNuevo Then
                    oCmdEjec.CommandText = "SP_REPARTIDORES_REGISTRAR"
                Else
                    oCmdEjec.CommandText = "SP_REPARTIDORES_MODIFICAR"
                End If

                On Error GoTo grabar

                Dim Smensaje As String

                Dim vIDr     As Integer

                Smensaje = ""
                vIDr = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@REPARTIDOR", adVarChar, adParamInput, 100, Trim(Me.txtRepartidor.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DNI", adChar, adParamInput, 8, Trim(Me.txtDni.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PLACAVEHICULO", adVarChar, adParamInput, 30, Trim(Me.txtPlaca.Text))

                If VNuevo Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adBigInt, adParamOutput, , vIDr)
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adBigInt, adParamInput, , Me.lblCodigo.Caption)
                End If
               
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 200, Smensaje)
                                
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DIRECCION", adVarChar, adParamInput, 150, Trim(Me.txtDireccion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FONO1", adVarChar, adParamInput, 30, Trim(Me.txtFono1.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FONO2", adVarChar, adParamInput, 30, Trim(Me.txtFono2.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MARCAVEHICULO", adVarChar, adParamInput, 30, Trim(Me.txtMarca.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MODELOVEHICULO", adVarChar, adParamInput, 30, Trim(Me.txtModelo.Text))
                
                oCmdEjec.Execute
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                
                Smensaje = oCmdEjec.Parameters("@EXITO").Value
                vIDr = oCmdEjec.Parameters("@IDREPARTIDOR").Value
                
                If Len(Trim(Smensaje)) = 0 Then
                    DesactivarControles Me
                
                    Estado_Botones grabar
                    Me.lvRepartidores.Enabled = True
                    Me.txtSearch.Enabled = True

                    If VNuevo Then

                        With Me.lvRepartidores.ListItems.Add(, , Trim(Me.txtRepartidor.Text))
                            .Tag = Trim(vIDr)
                            .SubItems(1) = Me.txtDireccion.Text
                            .SubItems(2) = Me.txtFono1.Text
                            .SubItems(3) = Me.txtFono2.Text
                            .SubItems(4) = "SI"
                        End With

                    Else
                        Me.lvRepartidores.SelectedItem.Text = Trim(Me.txtRepartidor.Text)
                        Me.lvRepartidores.SelectedItem.SubItems(1) = Me.txtDireccion.Text
                        Me.lvRepartidores.SelectedItem.SubItems(2) = Me.txtFono1.Text
                        Me.lvRepartidores.SelectedItem.SubItems(3) = Me.txtFono2.Text
                    End If

                    'set itemg=me.lvMesas.ListItems.Add(,,
                    MsgBox "Datos Almacenados Correctamente", vbInformation, Pub_Titulo
                Else
                    MsgBox Smensaje, vbInformation, Pub_Titulo
                End If

                Exit Sub

grabar:
                MsgBox Err.Description, vbInformation, NombreProyecto

            End If

        Case 3 'Modificar
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
    
            Me.txtRepartidor.SetFocus
            Me.txtSearch.Enabled = False

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvRepartidores.Enabled = True
            Me.txtSearch.Enabled = True
    
        Case 5 'Desactivar

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then

                Dim vELI As String

                vELI = ""

                On Error GoTo elimina
            
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "SP_REPARTIDORES_DESACT"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adBigInt, adParamInput, , Me.lblCodigo.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 200, vELI)
                oCmdEjec.Execute
                
                vELI = oCmdEjec.Parameters("@EXITO").Value
                
                If Len(Trim(vELI)) = 0 Then
                    LimpiarControles Me
                    
                    Me.lvRepartidores.Enabled = True
                    Me.lvRepartidores.SelectedItem.SubItems(4) = "NO"
                    Me.txtSearch.Enabled = True
                    Estado_Botones Eliminar
                    MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
                Else
                    MsgBox vELI, vbCritical, Pub_Titulo
                End If

                Exit Sub

elimina:
                MsgBox Err.Description, vbInformation, NombreProyecto
            End If

        Case 6 'ACTIVAR

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then

                Dim vACT As String

                vACT = ""

                On Error GoTo activa
            
                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "SP_REPARTIDORES_DESACT"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTIDOR", adBigInt, adParamInput, , Me.lblCodigo.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 200, vACT)
                oCmdEjec.Execute
                
                vACT = oCmdEjec.Parameters("@EXITO").Value
                
                If Len(Trim(vACT)) = 0 Then
                    LimpiarControles Me
                    
                    Me.lvRepartidores.Enabled = True
                    Me.lvRepartidores.SelectedItem.SubItems(4) = "SI"
                    Me.txtSearch.Enabled = True
                    Estado_Botones Eliminar
                    MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
                Else
                    MsgBox vACT, vbCritical, Pub_Titulo
                End If

                Exit Sub

activa:
                MsgBox Err.Description, vbInformation, NombreProyecto
            End If

    End Select

End Sub



Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.txtFono1.SetFocus
    Me.txtFono1.SelStart = 0
    Me.txtFono1.SelLength = Len(Me.txtFono1.Text)
End If
End Sub

Private Sub txtDni_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtDireccion.SetFocus
Me.txtDireccion.SelStart = 0
Me.txtDireccion.SelLength = Len(Me.txtDireccion.Text)
End If
End Sub

Private Sub txtFono1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Me.txtFono2.SetFocus
    Me.txtFono2.SelStart = 0
    Me.txtFono2.SelLength = Len(Me.txtFono2.Text)
End If
End Sub

Private Sub txtFono2_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtPlaca.SetFocus
Me.txtPlaca.SelStart = 0
Me.txtPlaca.SelLength = Len(Me.txtPlaca.Text)
End If
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtModelo.SetFocus
Me.txtModelo.SelStart = 0
Me.txtModelo.SelLength = Len(Me.txtModelo.Text)
End If
End Sub

Private Sub txtPlaca_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtMarca.SetFocus
Me.txtMarca.SelStart = 0
Me.txtMarca.SelLength = Len(Me.txtMarca.Text)
End If
End Sub

Private Sub txtRepartidor_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.txtDni.SetFocus
Me.txtDni.SelStart = 0
Me.txtDni.SelLength = Len(Me.txtDni.Text)
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_REPARTIDORES_LISTAR"
    oCmdEjec.CommandType = adCmdStoredProc
    Dim ORSd As ADODB.Recordset
    Set ORSd = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtSearch.Text))
    Me.lvRepartidores.ListItems.Clear
    Dim oITEM As Object
    Do While Not ORSd.EOF
        Set oITEM = Me.lvRepartidores.ListItems.Add(, , ORSd!REP)
        oITEM.Tag = ORSd!IDE
        oITEM.SubItems(1) = ORSd!dir
        oITEM.SubItems(2) = ORSd!f1
        oITEM.SubItems(3) = ORSd!F2
        oITEM.SubItems(4) = ORSd!ac
        ORSd.MoveNext
    Loop
End If
End Sub
