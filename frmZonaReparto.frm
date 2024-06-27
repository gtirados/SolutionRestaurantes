VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmZonaReparto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zonas de Reparto"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
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
   ScaleHeight     =   4650
   ScaleWidth      =   9780
   Begin TabDlg.SSTab SSTZona 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Zona de Reparto"
      TabPicture(0)   =   "frmZonaReparto.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "lblCodigo"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "txtDenominacion"
      Tab(0).Control(7)=   "txtTall"
      Tab(0).Control(8)=   "txtTarifa"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmZonaReparto.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lvZonas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView lvZonas 
         Height          =   3135
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5530
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
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   8415
      End
      Begin VB.TextBox txtTarifa 
         Height          =   285
         Left            =   -68640
         TabIndex        =   8
         Tag             =   "X"
         Top             =   2235
         Width           =   1455
      End
      Begin VB.TextBox txtTall 
         Height          =   285
         Left            =   -72120
         TabIndex        =   7
         Tag             =   "X"
         Top             =   2235
         Width           =   1575
      End
      Begin VB.TextBox txtDenominacion 
         Height          =   285
         Left            =   -72120
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "X"
         Top             =   1800
         Width           =   4935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "min"
         Height          =   195
         Left            =   -70520
         TabIndex        =   13
         Top             =   2280
         Width           =   315
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zona:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   510
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -72120
         TabIndex        =   9
         Tag             =   "X"
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tarifa S/."
         Height          =   195
         Left            =   -69480
         TabIndex        =   5
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo Aprox. de Llegada:"
         Height          =   435
         Left            =   -73680
         TabIndex        =   4
         Top             =   2160
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   -73560
         TabIndex        =   3
         Top             =   1845
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   195
         Left            =   -72840
         TabIndex        =   2
         Top             =   1365
         Width           =   675
      End
   End
   Begin MSComctlLib.ImageList iZona 
      Left            =   9840
      Top             =   5040
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
            Picture         =   "frmZonaReparto.frx":0038
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZonaReparto.frx":05D2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZonaReparto.frx":0B6C
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZonaReparto.frx":1106
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmZonaReparto.frx":16A0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbZona 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   635
      ButtonWidth     =   2143
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "iZona"
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
End
Attribute VB_Name = "frmZonaReparto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Sub Mandar_Datos()
With Me.lvZonas
Me.lblCodigo.Caption = .SelectedItem.Tag
    Me.txtDenominacion.Text = .SelectedItem.Text

    Me.txtTall.Text = .SelectedItem.SubItems(1)
    Me.txtTarifa.Text = .SelectedItem.SubItems(2)

    
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub ConfigurarLV()
With Me.lvZonas
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "Zona", 3500
    .ColumnHeaders.Add , , "Tiempo Aprox. Llegada", 1800
    .ColumnHeaders.Add , , "Tarifa"
    .ColumnHeaders.Add , , "Estado"
End With
End Sub

Private Sub ListarZonas()
Dim oRSr As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_ZONAREPARTO_LISTAR"
Set oRSr = oCmdEjec.Execute(, LK_CODCIA)

Do While Not oRSr.EOF

    With Me.lvZonas.ListItems.Add(, , Trim(oRSr!DENO))
        .Tag = Trim(oRSr!IDE)
        .SubItems(1) = oRSr!TL
        .SubItems(2) = oRSr!tf
        .SubItems(3) = oRSr!EST
    End With
   
    oRSr.MoveNext
Loop
End Sub

Private Sub Estado_Botones(val As Valores)
Select Case val
    Case InicializarFormulario, grabar, cancelar, Eliminar
        Me.tbZona.Buttons(1).Enabled = True
        Me.tbZona.Buttons(2).Enabled = False
        Me.tbZona.Buttons(3).Enabled = False
        Me.tbZona.Buttons(4).Enabled = False
        Me.tbZona.Buttons(5).Enabled = False
        Me.tbZona.Buttons(6).Enabled = False
        Me.SSTZona.tab = 0
    Case Nuevo, Editar
        Me.tbZona.Buttons(1).Enabled = False
        Me.tbZona.Buttons(2).Enabled = True
        Me.tbZona.Buttons(3).Enabled = False
        Me.tbZona.Buttons(4).Enabled = True
        Me.tbZona.Buttons(5).Enabled = False
         Me.tbZona.Buttons(6).Enabled = False
        Me.lvZonas.Enabled = False
        Me.txtSearch.Enabled = False
        Me.SSTZona.tab = 0
    Case buscar
        Me.tbZona.Buttons(1).Enabled = True
        Me.tbZona.Buttons(2).Enabled = False
        Me.tbZona.Buttons(3).Enabled = False
        Me.tbZona.Buttons(4).Enabled = False
        Me.SSTZona.tab = 1
    Case AntesDeActualizar
        Me.tbZona.Buttons(1).Enabled = False
        Me.tbZona.Buttons(2).Enabled = False
        Me.tbZona.Buttons(3).Enabled = True
        Me.tbZona.Buttons(4).Enabled = True
         Me.tbZona.Buttons(5).Enabled = True
          Me.tbZona.Buttons(6).Enabled = True
        Me.SSTZona.tab = 0
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
ListarZonas
End Sub

Private Sub lvZonas_DblClick()
Mandar_Datos
End Sub

Private Sub SSTZona_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
    txtSearch_KeyPress vbKeyReturn
End If
End Sub

Private Sub tbZona_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Estado_Botones Nuevo
            VNuevo = True
            Me.txtDenominacion.SetFocus

        Case 2 'Guardar
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtDenominacion.Text)) = 0 Then
                MsgBox "Debe ingresar el Código", vbCritical, NombreProyecto
                Me.txtDenominacion.SetFocus
            Else

                If VNuevo Then
                    oCmdEjec.CommandText = "SP_ZONAREPARTO_REGISTRAR"
                Else
                    oCmdEjec.CommandText = "SP_ZONAREPARTO_MODIFICAR"
                End If

                On Error GoTo grabar

                Dim Smensaje As String

                Dim vIDz     As Integer

                Smensaje = ""
                vIDz = 0

                oCmdEjec.Prepared = True
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DENOMINACION", adVarChar, adParamInput, 30, Trim(Me.txtDenominacion.Text))

                If Len(Trim(Me.txtTall.Text)) = 0 Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIEMPOLLEGADA", adInteger, adParamInput, , 0)
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIEMPOLLEGADA", adInteger, adParamInput, , Me.txtTall.Text)
                End If

                If Len(Trim(Me.txtTarifa.Text)) = 0 Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TARIFA", adDouble, adParamInput, , 0)
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TARIFA", adDouble, adParamInput, , CDbl(Me.txtTarifa.Text))
                End If

                If VNuevo Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adBigInt, adParamOutput, , vIDz)
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDZONA", adBigInt, adParamInput, , Me.lblCodigo.Caption)
                End If
                
                If Not VNuevo Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
                End If
 
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 200, Smensaje)

                oCmdEjec.Execute
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))

                Smensaje = oCmdEjec.Parameters("@EXITO").Value
                vIDz = oCmdEjec.Parameters("@IDZONA").Value

                If Len(Trim(Smensaje)) = 0 Then
                    DesactivarControles Me

                    Estado_Botones grabar
                    Me.lvZonas.Enabled = True
                    Me.txtSearch.Enabled = True

                    If VNuevo Then

                        With Me.lvZonas.ListItems.Add(, , Trim(Me.txtDenominacion.Text))
                            .Tag = Trim(vIDz)
                            .SubItems(1) = Me.txtTall.Text
                            .SubItems(2) = Me.txtTarifa.Text
                            .SubItems(3) = "SI"
                        End With

                    Else
                        Me.lvZonas.SelectedItem.Text = Trim(Me.txtDenominacion.Text)
                        Me.lvZonas.SelectedItem.SubItems(1) = Me.txtTall.Text
                        Me.lvZonas.SelectedItem.SubItems(2) = Me.txtTarifa.Text
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
    
            Me.txtDenominacion.SetFocus
            Me.txtSearch.Enabled = False

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.lvZonas.Enabled = True
            Me.txtSearch.Enabled = True
    
        Case 5 'Desactivar

            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then

                Dim vELI As String

                vELI = ""

                On Error GoTo elimina

                LimpiaParametros oCmdEjec
                oCmdEjec.Prepared = True
                oCmdEjec.CommandText = "SP_ZONAREPARTO_DESACT"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTO", adBigInt, adParamInput, , Me.lblCodigo.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , False)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 200, vELI)
                oCmdEjec.Execute

                vELI = oCmdEjec.Parameters("@EXITO").Value

                If Len(Trim(vELI)) = 0 Then
                    LimpiarControles Me

                    Me.lvZonas.Enabled = True
                    Me.lvZonas.SelectedItem.SubItems(3) = "NO"
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
                oCmdEjec.CommandText = "SP_ZONAREPARTO_DESACT"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDREPARTO", adBigInt, adParamInput, , Me.lblCodigo.Caption)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ESTADO", adBoolean, adParamInput, , True)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EXITO", adVarChar, adParamOutput, 200, vACT)
                oCmdEjec.Execute
            
                vACT = oCmdEjec.Parameters("@EXITO").Value
            
                If Len(Trim(vACT)) = 0 Then
                    LimpiarControles Me
            
                    Me.lvZonas.Enabled = True
                    Me.lvZonas.SelectedItem.SubItems(3) = "SI"
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

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_ZONAREPARTO_LISTAR"
    oCmdEjec.CommandType = adCmdStoredProc
    Dim ORSd As ADODB.Recordset
    Set ORSd = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtSearch.Text))
    Me.lvZonas.ListItems.Clear
    Dim oITEM As Object
    Do While Not ORSd.EOF
        Set oITEM = Me.lvZonas.ListItems.Add(, , ORSd!DENO)
        oITEM.Tag = ORSd!IDE
        oITEM.SubItems(1) = ORSd!TL
        oITEM.SubItems(2) = ORSd!tf
        oITEM.SubItems(3) = ORSd!EST
        ORSd.MoveNext
    Loop
End If
End Sub
