VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTurno 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Turno"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTurno.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9195
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
            Picture         =   "frmTurno.frx":08CA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTurno.frx":0E64
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTurno.frx":13FE
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTurno.frx":1998
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTurno.frx":1F32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMesa 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   635
      ButtonWidth     =   2037
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
   Begin TabDlg.SSTab stabMesa 
      Height          =   4095
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483636
      TabCaption(0)   =   "Turno"
      TabPicture(0)   =   "frmTurno.frx":262D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtAbreviatura"
      Tab(0).Control(1)=   "txtTolerancia"
      Tab(0).Control(2)=   "dtpIni"
      Tab(0).Control(3)=   "ComActivo"
      Tab(0).Control(4)=   "txtCodigo"
      Tab(0).Control(5)=   "txtDenominacion"
      Tab(0).Control(6)=   "dtpFin"
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(9)=   "Label7"
      Tab(0).Control(10)=   "Label6"
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(12)=   "Label3"
      Tab(0).Control(13)=   "Label1"
      Tab(0).Control(14)=   "Label2"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmTurno.frx":2649
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtBusMesa"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lvZonas"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.TextBox txtAbreviatura 
         Height          =   285
         Left            =   -70920
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "X"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtTolerancia 
         Height          =   285
         Left            =   -70920
         MaxLength       =   2
         TabIndex        =   5
         Tag             =   "X"
         Top             =   2880
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   300
         Left            =   -70920
         TabIndex        =   3
         Top             =   2040
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   160432131
         CurrentDate     =   41479
      End
      Begin VB.ComboBox ComActivo 
         Height          =   315
         ItemData        =   "frmTurno.frx":2665
         Left            =   -70920
         List            =   "frmTurno.frx":266F
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   -70920
         MaxLength       =   10
         TabIndex        =   0
         Tag             =   "X"
         Top             =   840
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvZonas 
         Height          =   3135
         Left            =   120
         TabIndex        =   12
         Top             =   840
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
      Begin VB.TextBox txtDenominacion 
         Height          =   285
         Left            =   -70920
         MaxLength       =   40
         TabIndex        =   1
         Tag             =   "X"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox txtBusMesa 
         Height          =   285
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   8055
      End
      Begin MSComCtl2.DTPicker dtpFin 
         Height          =   300
         Left            =   -70920
         TabIndex        =   4
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "hh:mm tt"
         Format          =   160432131
         CurrentDate     =   41479
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Abreviatura:"
         Height          =   195
         Left            =   -72120
         TabIndex        =   19
         Top             =   1680
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min."
         Height          =   195
         Left            =   -69720
         TabIndex        =   18
         Top             =   2925
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tolerancia:"
         Height          =   195
         Left            =   -72000
         TabIndex        =   17
         Top             =   2925
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Final:"
         Height          =   195
         Left            =   -71970
         TabIndex        =   16
         Top             =   2430
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Inicial:"
         Height          =   195
         Left            =   -72090
         TabIndex        =   15
         Top             =   2070
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Activo:"
         Height          =   195
         Left            =   -71640
         TabIndex        =   14
         Top             =   3420
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Left            =   -71715
         TabIndex        =   13
         Top             =   885
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Denominación:"
         Height          =   195
         Left            =   -72330
         TabIndex        =   9
         Top             =   1365
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private VNuevo As Boolean

Sub Mandar_Datos()
With Me.lvZonas
    Me.txtCodigo.Text = .SelectedItem.Text
    Me.txtDenominacion.Text = .SelectedItem.SubItems(1)
    Me.dtpIni.Value = .SelectedItem.SubItems(7)
    Me.dtpFin.Value = .SelectedItem.SubItems(8)
    Me.txtTolerancia.Text = .SelectedItem.SubItems(4)
    Me.ComActivo.ListIndex = .SelectedItem.SubItems(6)
    Me.txtabreviatura.Text = .SelectedItem.SubItems(9)
    Estado_Botones AntesDeActualizar
End With
End Sub

Private Sub ConfigurarLV()
With Me.lvZonas
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CODIGO"
    .ColumnHeaders.Add , , "DESCRIPCION", 4000
    .ColumnHeaders.Add , , "INICIO", 1000
    .ColumnHeaders.Add , , "FIN", 1000
    .ColumnHeaders.Add , , "TOLERANCIA", 1000
    .ColumnHeaders.Add , , "ACTIVO", 700
    .ColumnHeaders.Add , , "Valor", 0
    .ColumnHeaders.Add , , "", 0
    .ColumnHeaders.Add , , "", 0
    .ColumnHeaders.Add , , "ABRE", 0
End With
End Sub
Private Sub ListarZonas()
Dim oRsZonas As ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPLISTARTURNOS"

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)

Set oRsZonas = oCmdEjec.Execute

Dim tt As Object
Dim itemX As MSComctlLib.ListItem

Me.lvZonas.ListItems.Clear
Do While Not oRsZonas.EOF
    Set itemX = Me.lvZonas.ListItems.Add(, , oRsZonas!Codigo)
    itemX.SubItems(1) = oRsZonas!DESCRIPCION
    itemX.SubItems(2) = oRsZonas!ini
    itemX.SubItems(3) = oRsZonas!fin
    itemX.SubItems(4) = oRsZonas!TOLERANCIA
    itemX.SubItems(5) = oRsZonas!activo
    itemX.SubItems(6) = oRsZonas!valor
    itemX.SubItems(7) = oRsZonas!INIX
    itemX.SubItems(8) = oRsZonas!FINX
    itemX.SubItems(9) = oRsZonas!ABRE
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
CentrarFormulario MDIForm1, Me
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

Private Sub lvZonas_DblClick()
Mandar_Datos
End Sub

Private Sub tbMesa_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'NUEVO
            ActivarControles Me
            LimpiarControles Me
            Me.txtCodigo.Enabled = False
            Estado_Botones Nuevo
            VNuevo = True
            Me.ComActivo.ListIndex = 1
            Me.txtDenominacion.SetFocus

        Case 2 'Guardar
            LimpiaParametros oCmdEjec

            If Len(Trim(Me.txtDenominacion.Text)) = 0 Then
                MsgBox "Debe ingresar la Denominación del turno", vbCritical, Pub_Titulo
                Me.txtDenominacion.SetFocus
                ElseIf Len(Trim(Me.txtabreviatura.Text)) = 0 Then
                MsgBox "Debe ingresar la Abreviatura del Turno", vbCritical, Pub_Titulo
                Me.txtabreviatura.SetFocus
            ElseIf Not IsNumeric(Me.txtTolerancia.Text) Then
                MsgBox "La Tolerancia ingresada no es correcta.", vbCritical, Pub_Titulo
                Me.txtTolerancia.SetFocus
                Me.txtTolerancia.SelStart = 0
                Me.txtTolerancia.SelLength = Len(Me.txtTolerancia.Text)
            ElseIf val(Me.txtTolerancia.Text) > 60 Then
                MsgBox "La Tolerancia supera los 60 minutos.", vbCritical, Pub_Titulo
                Me.txtTolerancia.SetFocus
                Me.txtTolerancia.SelStart = 0
                Me.txtTolerancia.SelLength = Len(Me.txtTolerancia.Text)

                Exit Sub
   
            Else
    
                On Error GoTo grabar

                Dim vCodigo As Integer

                oCmdEjec.Prepared = True
        
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DESCRIPCION", adVarChar, adParamInput, 40, Trim(Me.txtDenominacion.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ABREVIATURA", adVarChar, adParamInput, 2, Trim(Me.txtabreviatura.Text))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@INI", adVarChar, adParamInput, 8, Format(Me.dtpIni.Value, "hh:mm:ss"))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FIN", adVarChar, adParamInput, 8, Format(Me.dtpFin.Value, "hh:mm:ss"))
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOLERANCIA", adInteger, adParamInput, , Me.txtTolerancia.Text)
        
                If VNuevo Then
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamOutput, , vCodigo)
                    oCmdEjec.CommandText = "SPREGISTRARTURNO"
                Else
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , Me.txtCodigo.Text)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Activo", adInteger, adParamInput, , Me.ComActivo.ListIndex)
                    oCmdEjec.CommandText = "SPMODIFICARTURNO"
                End If
        
                oCmdEjec.Execute
                vCodigo = oCmdEjec.Parameters("@Codigo").Value
        
                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
                DesactivarControles Me
                Estado_Botones grabar
                ListarZonas
                Me.lvZonas.Enabled = True
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
