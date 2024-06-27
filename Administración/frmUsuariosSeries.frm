VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUsuariosSeries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Series por Usuarios"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
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
   ScaleHeight     =   4815
   ScaleWidth      =   8235
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   635
      ButtonWidth     =   2037
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmUsuariosSeries.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblNameTipo"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "lblUsuario"
      Tab(0).Control(4)=   "lvDatos"
      Tab(0).Control(5)=   "DatDocumento"
      Tab(0).Control(6)=   "txtSerie"
      Tab(0).Control(7)=   "cmdDel"
      Tab(0).Control(8)=   "cmdAdd"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmUsuariosSeries.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtSearch"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lvUsuarios"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ListView lvUsuarios 
         Height          =   3135
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5530
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
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   5655
      End
      Begin VB.CommandButton cmdAdd 
         Enabled         =   0   'False
         Height          =   360
         Left            =   -68400
         TabIndex        =   9
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton cmdDel 
         Enabled         =   0   'False
         Height          =   360
         Left            =   -68400
         TabIndex        =   8
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txtSerie 
         Enabled         =   0   'False
         Height          =   285
         Left            =   -70320
         TabIndex        =   5
         Top             =   1215
         Width           =   1215
      End
      Begin MSDataListLib.DataCombo DatDocumento 
         Height          =   315
         Left            =   -73680
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin MSComctlLib.ListView lvDatos 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   6
         Top             =   1680
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4048
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -73680
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie:"
         Height          =   195
         Left            =   -70920
         TabIndex        =   3
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documento:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   2
         Top             =   1260
         Width           =   1050
      End
      Begin VB.Label lblNameTipo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   -74760
         TabIndex        =   1
         Top             =   765
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmUsuariosSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private gTipoFacturacion As String

Private Sub LlenarUsuarios()

    Dim orsU  As New ADODB.Recordset

    Dim itemU As Object

    LimpiaParametros oCmdEjec

    If gTipoFacturacion = "U" Then
        oCmdEjec.CommandText = "SP_USUARIOS_LIST"
         Set orsU = oCmdEjec.Execute

    Do While Not orsU.EOF
        Set itemU = Me.lvUsuarios.ListItems.Add(, , orsU!usuario)
        itemU.Tag = Trim(orsU!USU_KEY)
        orsU.MoveNext
    Loop
    
    ElseIf gTipoFacturacion = "A" Then
        oCmdEjec.CommandText = "SP_PARGEN_LIST"
         Set orsU = oCmdEjec.Execute

    Do While Not orsU.EOF
        Set itemU = Me.lvUsuarios.ListItems.Add(, , orsU!PG)
        itemU.Tag = Trim(orsU!PAR_CODCIA)
        orsU.MoveNext
    Loop
    
    End If
   
   

End Sub

Private Sub ConfigurarLV()
With Me.lvDatos
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "DOCUMENTO"
    .ColumnHeaders.Add , , "SERIE", 600
End With

With Me.lvUsuarios
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "USUARIO", 4000
End With
End Sub

Private Sub cmdAdd_Click()

    If Me.DatDocumento.BoundText = "" Then
        MsgBox "Debe elegir el documento.", vbCritical, Pub_Titulo
        Me.DatDocumento.SetFocus
    ElseIf Len(Trim(Me.txtSerie.Text)) = 0 Then
        MsgBox "Debe ingresar la serie.", vbCritical, Pub_Titulo
    Else

        Dim itemT As Object
        Dim xDato As Boolean
        xDato = False
        If Me.lvDatos.ListItems.count = 0 Then
            Set itemT = Me.lvDatos.ListItems.Add(, , Me.DatDocumento.Text)
            itemT.Tag = Me.DatDocumento.BoundText
            itemT.SubItems(1) = Me.txtSerie.Text
            
            Me.DatDocumento.BoundText = ""
            Me.txtSerie.Text = ""
        Else

            For Each itemT In Me.lvDatos.ListItems

                If itemT.Tag = Me.DatDocumento.BoundText Then
                    xDato = True
                    Exit For
                End If

            Next
            If xDato Then
                MsgBox "Documento ya agregado.", vbInformation, Pub_Titulo
            Else
                Set itemT = Me.lvDatos.ListItems.Add(, , Me.DatDocumento.Text)
                itemT.Tag = Me.DatDocumento.BoundText
                itemT.SubItems(1) = Me.txtSerie.Text
            End If
            Me.DatDocumento.BoundText = ""
            Me.txtSerie.Text = ""
        End If
    
    End If

End Sub

Private Sub cmdDel_Click()
Me.lvDatos.ListItems.Remove Me.lvDatos.SelectedItem.Index
End Sub

Private Sub Form_Load()

    'VERIFICANDO EL TIPO DE FACTURACION
    Dim ORStf As New ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_FACTURACION_TIPO"
    Set ORStf = oCmdEjec.Execute(, LK_CODCIA)

    gTipoFacturacion = ORStf!tipo

    ConfigurarLV

    Dim ORStd As New ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_TIPOS_DOCTOS_LIST"
    Set ORStd = oCmdEjec.Execute
    Set Me.DatDocumento.RowSource = ORStd
    Me.DatDocumento.ListField = "NOMBRE"
    Me.DatDocumento.BoundColumn = "CODIGO"
        
    If gTipoFacturacion = "U" Then
       
        Me.lblNameTipo.Caption = "Usuario:"
    ElseIf gTipoFacturacion = "A" Then
        Me.lblNameTipo.Caption = "Compañía:"
    End If

    LlenarUsuarios
    Me.Toolbar1.Buttons(1).Enabled = False
    Me.Toolbar1.Buttons(2).Enabled = False
    Me.Toolbar1.Buttons(3).Enabled = False
End Sub

Private Sub lvUsuarios_DblClick()
    Me.lblUsuario.Caption = Me.lvUsuarios.SelectedItem.Text
    Me.lblUsuario.Tag = Me.lvUsuarios.SelectedItem.Tag
    LimpiaParametros oCmdEjec

    If gTipoFacturacion = "U" Then
        oCmdEjec.CommandText = "SP_USUARIOS_SERIES_FILL"
    ElseIf gTipoFacturacion = "A" Then
        oCmdEjec.CommandText = "SP_PARGEN_SERIES_FILL"
    End If

    Dim ORSf As New ADODB.Recordset

    Set ORSf = oCmdEjec.Execute(, Me.lvUsuarios.SelectedItem.Tag)

    Dim ITEMf As Object

    Do While Not ORSf.EOF
        Set ITEMf = Me.lvDatos.ListItems.Add(, , ORSf!TD)
        ITEMf.Tag = ORSf!cod
        ITEMf.SubItems(1) = ORSf!serie
        ORSf.MoveNext
    Loop

    Me.SSTab1.tab = 0
    Me.Toolbar1.Buttons(1).Enabled = True
    Me.Toolbar1.Buttons(3).Enabled = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'modiicar
            VNuevo = False
            Me.Toolbar1.Buttons(1).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = True
            Me.Toolbar1.Buttons(3).Enabled = True
            ActivarControles Me
            Me.DatDocumento.SetFocus
            Me.DatDocumento.Enabled = True
            Me.cmdAdd.Enabled = True
            Me.cmdDel.Enabled = True
            Me.lvDatos.Enabled = True

        Case 2 'grbar
        
            On Error GoTo grabar

            Pub_ConnAdo.BeginTrans
            LimpiaParametros oCmdEjec
   Dim ITEMd As Object
            If gTipoFacturacion = "U" Then
               
                oCmdEjec.CommandText = "SP_USUARIOS_SERIES_DELETE"
                oCmdEjec.Execute , Me.lblUsuario.Tag
            
                oCmdEjec.CommandText = "SP_USUARIOS_SERIES_REGISTRAR"

             

                For Each ITEMd In Me.lvDatos.ListItems

                    LimpiaParametros oCmdEjec
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, Me.lblUsuario.Tag)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPODOCTO", adChar, adParamInput, 2, ITEMd.Tag)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamInput, 3, ITEMd.SubItems(1))
                    oCmdEjec.Execute
                Next

            Else
                oCmdEjec.CommandText = "SP_PARGEN_SERIES_DELETE"
                oCmdEjec.Execute , LK_CODCIA
            
                oCmdEjec.CommandText = "SP_PARGEN_SERIES_REGISTRAR"

                

                For Each ITEMd In Me.lvDatos.ListItems

                    LimpiaParametros oCmdEjec
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, Me.lblUsuario.Tag)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIPODOCTO", adChar, adParamInput, 2, ITEMd.Tag)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SERIE", adChar, adParamInput, 3, ITEMd.SubItems(1))
                    oCmdEjec.Execute
                Next

            End If

            Me.lvDatos.ListItems.Clear
            Me.DatDocumento.BoundText = ""
            Me.txtSerie.Text = ""
            Me.lblUsuario.Caption = ""
            Me.lblUsuario.Tag = ""
            Me.SSTab1.tab = 1
            Me.Toolbar1.Buttons(1).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(3).Enabled = False
            Pub_ConnAdo.CommitTrans
            
            Exit Sub

grabar:
            Pub_ConnAdo.RollbackTrans
            MsgBox Err.Description, vbInformation, NombreProyecto

            '''
            '''            If Me.lvDatos.ListItems.count = 0 Then
            '''                MsgBox "Debe agregar documentos al usuario", vbCritical, Pub_Titulo
            '''
            '''            Else
            '''
            '''                On Error GoTo grabar
            '''
            '''                Dim vCodigo As Integer
            '''
            '''                oCmdEjec.Prepared = True
            '''
            '''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODIGO", adChar, adParamInput, 2, Me.txtCodigo.Text)
            '''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NOMBRE", adVarChar, adParamInput, 80, Trim(Me.txtDeno.Text))
            '''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ACTIVO", adBoolean, adParamInput, , Me.ComActivo.ListIndex)
            '''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DEFECTO", adBoolean, adParamInput, , Me.comDefecto.ListIndex)
            '''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@EDITABLE", adBoolean, adParamInput, , Me.comEditable.ListIndex)
            '''                'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TOLERANCIA", adInteger, adParamInput, , Me.txtTolerancia.Text)
            '''
            '''                If VNuevo Then
            '''                    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamOutput, , vCodigo)
            '''                    oCmdEjec.CommandText = "SP_TIPOSDOCTOS_REGISTRAR"
            '''                Else
            '''                    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , Me.txtCodigo.Text)
            '''                    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Activo", adInteger, adParamInput, , Me.ComActivo.ListIndex)
            '''                    oCmdEjec.CommandText = "SP_TIPOSDOCTOS_MODIFICAR"
            '''                End If
            '''
            '''                oCmdEjec.Execute
            '''                ' vCodigo = oCmdEjec.Parameters("@Codigo").Value
            '''
            '''                'oCmdEjec.Execute , Array(CodCia, Trim(Me.txtCodigo.Text), Trim(Me.txtDenominacion.Text), Trim(Me.txtZona.Text))
            '''                DesactivarControles Me
            '''                Estado_Botones grabar
            '''                ListarTD
            '''                Me.lvDatos.Enabled = True
            '''                Me.txtSearch.Enabled = True
            '''
            '''                'set itemg=me.lvMesas.ListItems.Add(,,
            '''                MsgBox "Datos Almacenados Correctamente", vbInformation, NombreProyecto
            '''
            '''                Exit Sub
            '''
            '''grabar:
            '''                MsgBox Err.Description, vbInformation, NombreProyecto
            '''
            '''            End If

        Case 3 'Cancelar
            Me.lvDatos.ListItems.Clear
            Me.DatDocumento.BoundText = ""
            Me.txtSerie.Text = ""
            Me.lblUsuario.Caption = ""
            Me.lblUsuario.Tag = ""
            Me.SSTab1.tab = 1
            Me.Toolbar1.Buttons(1).Enabled = False
            Me.Toolbar1.Buttons(2).Enabled = False
            Me.Toolbar1.Buttons(3).Enabled = False

    End Select

End Sub
