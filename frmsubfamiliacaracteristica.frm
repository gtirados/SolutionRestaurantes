VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmsubfamiliacaracteristica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caracteristicas de SubFamilias"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9030
   Begin MSComctlLib.Toolbar tbSF 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   741
      ButtonWidth     =   2037
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "ilMesa"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Modificar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab stabSF 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   8070
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
      TabCaption(0)   =   "Caracteristica"
      TabPicture(0)   =   "frmsubfamiliacaracteristica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblSubFamilia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvCaracteristicas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtCaracteristica"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAdd"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdDel"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdEdit"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Listado"
      TabPicture(1)   =   "frmsubfamiliacaracteristica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvSF"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdEdit 
         Height          =   360
         Left            =   7920
         Picture         =   "frmsubfamiliacaracteristica.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2520
         Width           =   510
      End
      Begin VB.CommandButton cmdDel 
         Height          =   360
         Left            =   7440
         Picture         =   "frmsubfamiliacaracteristica.frx":03C2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   360
         Left            =   6960
         Picture         =   "frmsubfamiliacaracteristica.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtCaracteristica 
         Height          =   285
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   7
         Tag             =   "X"
         Top             =   1440
         Width           =   4695
      End
      Begin MSComctlLib.ListView lvSF 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   7011
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
      Begin MSComctlLib.ListView lvCaracteristicas 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   2280
         Width           =   7815
         _ExtentX        =   13785
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
      Begin VB.Label lblSubFamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   5955
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caracteristica:"
         Height          =   195
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SubFamilia:"
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   1005
      End
   End
   Begin MSComctlLib.ImageList ilMesa 
      Left            =   9000
      Top             =   1320
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
            Picture         =   "frmsubfamiliacaracteristica.frx":0AD6
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubfamiliacaracteristica.frx":1070
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubfamiliacaracteristica.frx":160A
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubfamiliacaracteristica.frx":1BA4
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmsubfamiliacaracteristica.frx":213E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmsubfamiliacaracteristica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oRSeliminados As ADODB.Recordset

Private Sub CrearRecordsets()
 Set oRSeliminados = New ADODB.Recordset
    
    oRSeliminados.Fields.Append "IDSUBFAMILIA", adBigInt
    oRSeliminados.Fields.Append "IDCARACTERISTICA", adBigInt

    oRSeliminados.CursorLocation = adUseClient
    oRSeliminados.LockType = adLockOptimistic
    oRSeliminados.CursorType = adOpenDynamic
    oRSeliminados.Open
End Sub
Private Sub Estado_Botones(val As Valores)

    Select Case val

        Case InicializarFormulario, grabar, cancelar, Eliminar
            Me.tbSF.Buttons(1).Enabled = True
            Me.tbSF.Buttons(2).Enabled = False
            Me.tbSF.Buttons(3).Enabled = False
            Me.tbSF.Buttons(4).Enabled = False
            'Me.tbSF.Buttons(5).Enabled = False
            Me.cmdAdd.Enabled = False
            Me.cmdDel.Enabled = False
            Me.cmdEdit.Enabled = False
            Me.stabSF.tab = 0

        Case Nuevo, Editar
            Me.tbSF.Buttons(1).Enabled = False
            Me.tbSF.Buttons(2).Enabled = True
            Me.tbSF.Buttons(3).Enabled = False
            Me.tbSF.Buttons(4).Enabled = True
            Me.cmdAdd.Enabled = True
            Me.cmdDel.Enabled = True
            Me.cmdEdit.Enabled = True
            'Me.tbSF.Buttons(5).Enabled = False
            '        Me.lvZonas.Enabled = False
            '        Me.txtBusMesa.Enabled = False
            Me.stabSF.tab = 0

        Case buscar
            Me.tbSF.Buttons(1).Enabled = True
            Me.tbSF.Buttons(2).Enabled = False
            Me.tbSF.Buttons(3).Enabled = False
            Me.tbSF.Buttons(4).Enabled = False
            Me.stabSF.tab = 1

        Case AntesDeActualizar
            Me.tbSF.Buttons(1).Enabled = False
            Me.tbSF.Buttons(2).Enabled = False
            Me.tbSF.Buttons(3).Enabled = True
            Me.tbSF.Buttons(4).Enabled = True
            'Me.tbSF.Buttons(5).Enabled = True
            Me.stabSF.tab = 0
    End Select

End Sub

Private Sub cmdAdd_Click()

    If Me.lvCaracteristicas.ListItems.count = 0 Then

        Dim i As Object

        Set i = Me.lvCaracteristicas.ListItems.Add(, , Me.txtCaracteristica.Text)
        i.Tag = 0
        Me.txtCaracteristica.Text = ""
    Else

        Dim v As Boolean

        v = False

        Dim x As Object

        For Each x In Me.lvCaracteristicas.ListItems

            If x.Text = Me.txtCaracteristica.Text Then
                v = True

                Exit For

            End If

        Next

        If v Then
            MsgBox "Dato repetido", vbCritical, Pub_Titulo

            Exit Sub

        End If

        Set x = Me.lvCaracteristicas.ListItems.Add(, , Me.txtCaracteristica.Text)
        x.Tag = 0
        Me.txtCaracteristica.Text = ""
    End If

End Sub

Private Sub cmdDel_Click()

    If Me.lvCaracteristicas.SelectedItem.Tag <> 0 Then
        oRSeliminados.AddNew
        oRSeliminados!IDSUBFAMILIA = Me.lblSubFamilia.Tag
        oRSeliminados!IDCARACTERISTICA = Me.lvCaracteristicas.SelectedItem.Tag
        oRSeliminados.Update
    End If
Me.lvCaracteristicas.ListItems.Remove Me.lvCaracteristicas.SelectedItem.Index

End Sub

Private Sub cmdEdit_Click()
    frmsubfamiliacaracteristica_edit.txtCaracteristica.Text = Me.lvCaracteristicas.SelectedItem.Text
    frmsubfamiliacaracteristica_edit.Show vbModal

    If frmsubfamiliacaracteristica_edit.gAcepta Then
        Me.lvCaracteristicas.SelectedItem.Text = frmsubfamiliacaracteristica_edit.gCarac
    End If

End Sub

Private Sub Form_Load()
Estado_Botones InicializarFormulario
DesactivarControles Me
ConfigurarLV
CargarSubFamilias
Me.stabSF.tab = 1
CrearRecordsets
End Sub

Private Sub ConfigurarLV()
With Me.lvSF
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CODIGO"
    .ColumnHeaders.Add , , "SUBFAMILIA", 4000
End With
With Me.lvCaracteristicas
    .Gridlines = True
    .LabelEdit = lvwManual
    .View = lvwReport
    .FullRowSelect = True
    .ColumnHeaders.Add , , "CARACTERISTICA", 6000
End With
End Sub

Private Sub CargarSubFamilias()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CARACTERISTICA_SUBFAMILIA_LIST"
    oCmdEjec.CommandType = adCmdStoredProc

    Dim oRS As ADODB.Recordset

    Set oRS = oCmdEjec.Execute(, LK_CODCIA)

    Dim item As Object

    Do While Not oRS.EOF
        Set item = Me.lvSF.ListItems.Add(, , oRS!cod)
        item.SubItems(1) = oRS!nom
        oRS.MoveNext
    Loop

End Sub

Private Sub lvSF_DblClick()
    TraerCaracteristicas Me.lvSF.SelectedItem.Text
    
    Me.lblSubFamilia.Caption = Me.lvSF.SelectedItem.SubItems(1)
    Me.lblSubFamilia.Tag = Me.lvSF.SelectedItem.Text
    Estado_Botones AntesDeActualizar
End Sub

Private Sub tbSF_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

            '        Case 1 'NUEVO
            '            ActivarControles Me
            '            LimpiarControles Me
            '            Me.txtCodigo.Enabled = False
            '            Estado_Botones Nuevo
            '            VNuevo = True
            '            Me.ComActivo.ListIndex = 1
            '            Me.txtDenominacion.SetFocus

        Case 2 'Guardar
            

          

                On Error GoTo grabar

                Dim vCodigo As Integer

                Pub_ConnAdo.BeginTrans
                oCmdEjec.Prepared = True
                
                Dim idet As Object
                
                For Each idet In Me.lvCaracteristicas.ListItems
                
                    LimpiaParametros oCmdEjec

                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDSUBFAMILIA", adBigInt, adParamInput, , Me.lblSubFamilia.Tag)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CARACTERISTICA", adVarChar, adParamInput, 30, Trim(idet.Text))
                        
                    If idet.Tag <> 0 Then 'modifica
                        oCmdEjec.CommandText = "SP_CARACTERISTICA_MODIFICAR"
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCARACTERISTICA", adBigInt, adParamInput, , Trim(idet.Tag))
                    Else 'NUEVO
                        oCmdEjec.CommandText = "SP_CARACTERISTICA_REGISTRAR"
                    End If

                    oCmdEjec.Execute
                Next
                If oRSeliminados.RecordCount <> 0 Then
                    oRSeliminados.MoveFirst
                End If
                Do While Not oRSeliminados.EOF
                    LimpiaParametros oCmdEjec
                    oCmdEjec.CommandText = "SP_CARACTERISTICA_ELIMINAR"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDSUBFAMILIA", adBigInt, adParamInput, , oRSeliminados!IDSUBFAMILIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCARACTERISTICA", adBigInt, adParamInput, , oRSeliminados!IDCARACTERISTICA)
                    oCmdEjec.Execute
                    oRSeliminados.MoveNext
                
                Loop
                
                Pub_ConnAdo.CommitTrans
                MsgBox "Datos Almacenados Correctamente", vbInformation, Pub_Titulo
                Estado_Botones grabar
                DesactivarControles Me
                Me.lvCaracteristicas.ListItems.Clear
                Me.stabSF.tab = 1
                
                If oRSeliminados.RecordCount <> 0 Then
                oRSeliminados.MoveFirst
    Do While Not oRSeliminados.EOF
        oRSeliminados.Delete adAffectCurrent
        oRSeliminados.MoveNext
    Loop
                End If

                Exit Sub

grabar:
                Pub_ConnAdo.RollbackTrans
                MsgBox Err.Description, vbInformation, NombreProyecto

           

        Case 3 'Modificar
            VNuevo = False
            Estado_Botones Editar
            ActivarControles Me
            '            Me.txtCodigo.Enabled = False
            '            Me.ComActivo.Enabled = True
            '            Me.txtDenominacion.SetFocus

        Case 4 'Cancelar
            Estado_Botones cancelar
            DesactivarControles Me
            Me.stabSF.tab = 1
            '            Me.lvZonas.Enabled = True
            '            Me.txtBusMesa.Enabled = True

        Case 5 'Eliminar

            '            If MsgBox("¿Desea continuar con la Operación?", vbQuestion + vbYesNo, NombreProyecto) = vbYes Then
            '
            '                On Error GoTo elimina
            '
            '                LimpiaParametros oCmdEjec
            '                oCmdEjec.Prepared = True
            '                oCmdEjec.CommandText = "SpEliminarZona"
            '                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            '                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Codigo", adInteger, adParamInput, , CInt(Me.txtCodigo.Text))
            '                'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodZon", adInteger, adParamInput, , CInt(Me.dcboZona.BoundText))
            '                oCmdEjec.Execute
            '                LimpiarControles Me
            '                Me.lvZonas.Enabled = True
            '                Me.lvZonas.ListItems.Remove Me.lvZonas.SelectedItem.Index
            '                Me.txtBusMesa.Enabled = True
            '                Estado_Botones Eliminar
            '
            '                Exit Sub
            '
            'elimina:
            '                MsgBox Err.Description, vbInformation, NombreProyecto
            '            End If

    End Select

End Sub

Private Sub TraerCaracteristicas(xIDsubfamilia As Integer)
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SP_CARACTERISTICA_SUBFAMILIA_FILL"
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDSUBFAMILIA", adBigInt, adParamInput, , xIDsubfamilia)

    Dim ORSd As ADODB.Recordset

    Me.lvCaracteristicas.ListItems.Clear
    Set ORSd = oCmdEjec.Execute

    Dim itemd As Object

    Do While Not ORSd.EOF
        Set itemd = Me.lvCaracteristicas.ListItems.Add(, , ORSd!caracteristica)
        itemd.Tag = ORSd!IDCARACTERISTICA
        ORSd.MoveNext
    Loop

End Sub
