VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProgramacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programación"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
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
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuitar 
      Height          =   360
      Left            =   8280
      Picture         =   "frmProgramacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox txtCantidad 
      Height          =   285
      Left            =   7200
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtProducto 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   6975
   End
   Begin MSDataListLib.DataCombo DatTurno 
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   547
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   94830593
      CurrentDate     =   41479
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9120
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProgramacion.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProgramacion.frx":0724
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   635
      ButtonWidth     =   2011
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Grabar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvProductos 
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvDetalle 
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3413
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
   Begin VB.CommandButton cmdAgregar 
      Height          =   360
      Left            =   8280
      Picture         =   "frmProgramacion.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cant."
      Height          =   195
      Left            =   7320
      TabIndex        =   13
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Producto:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label lblFin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4800
      TabIndex        =   8
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label lblIni 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fin:"
      Height          =   195
      Left            =   4320
      TabIndex        =   6
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio:"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   1260
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Turno:"
      Height          =   195
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   570
   End
End
Attribute VB_Name = "frmProgramacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vBuscar   As Boolean 'variable para la busqueda de Productos

Private loc_key   As Integer
Public VNuevo As Boolean
Public xTurno As Integer
Private xIDp() As Double
Public xIDPROGRAMACION As Integer
Private Xim As Integer

Private Sub cmdAgregar_Click()

    If Len(Trim(Me.txtProducto.Tag)) = 0 Then
        MsgBox "Debe buscar el Producto.", vbCritical, Pub_Titulo
        Me.txtProducto.SetFocus

        Exit Sub

    End If

    If Not IsNumeric(Me.txtCantidad.Text) Then
        MsgBox "La Cantidad proporcionada es incorrecta.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

        Exit Sub

    End If

    If val(Me.txtCantidad.Text) <= 0 Then
        MsgBox "La Cantidad no puede ser menor o igual a Cero.", vbCritical, Pub_Titulo
        Me.txtCantidad.SetFocus
        Me.txtCantidad.SelStart = 0
        Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)

        Exit Sub

    End If

    Dim vAgrega As Boolean

    vAgrega = True

    Dim C As Integer

    For C = 1 To Me.lvDetalle.ListItems.count

        If Me.lvDetalle.ListItems(C).Tag = Me.txtProducto.Tag Then
            vAgrega = False

            Exit For

        End If

    Next

    If vAgrega Then

        Dim ItemP As Object

        Set ItemP = Me.lvDetalle.ListItems.Add(, , Me.txtProducto.Text)
        ItemP.Tag = Me.txtProducto.Tag
        ItemP.SubItems(1) = Me.txtCantidad.Text
    
        Me.txtProducto.Tag = ""
        Me.txtProducto.Text = ""
        Me.txtCantidad.Text = ""
        Me.txtProducto.SetFocus
    Else
        MsgBox "El producto ya fue agregado.", vbCritical, Pub_Titulo

        Me.txtProducto.SetFocus
        Me.txtProducto.SelStart = 0
        Me.txtProducto.SelLength = Len(Me.txtProducto.Text)
    End If

End Sub

Private Sub cmdQuitar_Click()

    If Me.lvDetalle.ListItems.count = 0 Then Exit Sub
    
    If Not VNuevo Then
        ReDim Preserve xIDp(Xim)
        xIDp(Xim) = Me.lvDetalle.SelectedItem.Tag
        Xim = Xim + 1
    End If

    Me.lvDetalle.ListItems.Remove Me.lvDetalle.SelectedItem.Index
End Sub

Private Sub DatTurno_Change()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPTURNO_OBTENERHORARIO"

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adInteger, adParamInput, , Me.DatTurno.BoundText)

Dim ORSh As ADODB.Recordset
Set ORSh = oCmdEjec.Execute

If Not ORSh.EOF Then
    Me.lblIni.Caption = ORSh!ini
    Me.lblFin.Caption = ORSh!fin
End If

End Sub





Private Sub DatTurno_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.txtProducto.SetFocus
End Sub

Private Sub dtpFecha_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Me.DatTurno.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    ListarTurno
    Xim = 0
    ReDim xIDp(0)
    Me.dtpFecha.Value = frmProgramacionOpcion.gFECHA
    ConfiguraLVs

    If Not VNuevo Then

        Dim orsP As ADODB.Recordset

        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPPROGRAMACION_FILL"
 
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , xIDPROGRAMACION)

        Set orsP = oCmdEjec.Execute

        If Not orsP.EOF Then
        Me.dtpFecha.Value = orsP!fecha
        Me.DatTurno.BoundText = orsP!IDETURNO
        
        Dim ItemP As Object
        Do While Not orsP.EOF
            Set ItemP = Me.lvDetalle.ListItems.Add(, , orsP!producto)
            ItemP.Tag = orsP!IDEPRODUCTO
            ItemP.SubItems(1) = orsP!Cantidad
            orsP.MoveNext
        Loop
        End If
    End If

End Sub


Private Sub ListarTurno()
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SPLISTARTURNOS"

oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@LISTADO", adBoolean, adParamInput, , 1)

Dim ORSt As ADODB.Recordset
Set ORSt = oCmdEjec.Execute

Set Me.DatTurno.RowSource = ORSt
Me.DatTurno.ListField = ORSt.Fields(1).Name
Me.DatTurno.BoundColumn = ORSt.Fields(0).Name
Me.DatTurno.BoundText = xTurno
End Sub

Private Sub ConfiguraLVs()

  

    With Me.lvDetalle
        .ColumnHeaders.Add , , "Descripción", 6000
        
        .ColumnHeaders.Add , , "Cant.", 800
        
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
        .MultiSelect = False
        
    End With
    
    With Me.lvProductos
        .ColumnHeaders.Add , , "PRODUCTO", 4000
        .Gridlines = True
        .LabelEdit = lvwManual
        .FullRowSelect = True
        .View = lvwReport
        .MultiSelect = False
    
    End With

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1 'GRABAR
          
            'AQUI VALIDACIONES
            Dim vVAL As Boolean

'''            If VNuevo Then
'''                LimpiaParametros oCmdEjec
'''                oCmdEjec.CommandText = "SPPROGRAMACION_VALIDAREPETIDO"
'''
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adBigInt, adParamInput, , Me.DatTurno.BoundText)
'''                'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , xIDPROGRAMACION)
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DATO", adBoolean, adParamOutput, , vVAL)
'''
'''                oCmdEjec.Execute
'''
'''                vVAL = oCmdEjec.Parameters("@DATO").Value
'''
'''                If vVAL Then
'''                    MsgBox "REPETIDO", vbCritical, Pub_Titulo
'''
'''                    Exit Sub
'''
'''                End If
'''
'''            Else
'''                LimpiaParametros oCmdEjec
'''                oCmdEjec.CommandText = "SPPROGRAMACION_VALIDAREPETIDO2"
'''
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adBigInt, adParamInput, , Me.DatTurno.BoundText)
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , xIDPROGRAMACION)
'''                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@DATO", adBoolean, adParamOutput, , vVAL)
'''
'''                oCmdEjec.Execute
'''
'''                vVAL = oCmdEjec.Parameters("@DATO").Value
'''
'''                If vVAL Then
'''                    MsgBox "REPETIDO", vbCritical, Pub_Titulo
'''
'''                    Exit Sub
'''
'''                End If
'''
'''            End If

            LimpiaParametros oCmdEjec

            On Error GoTo Graba

            Pub_ConnAdo.BeginTrans

            If VNuevo Then

                Dim vIDP As Integer

                oCmdEjec.CommandText = "SPPROGRAMACION_REGISTRAR"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adInteger, adParamInput, , Me.DatTurno.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamOutput, , vIDP)
                oCmdEjec.Execute
                vIDP = oCmdEjec.Parameters("@IDPROGRAMACION").Value
                
                Dim ip As Integer
                
                For ip = 1 To Me.lvDetalle.ListItems.count
                    LimpiaParametros oCmdEjec
                
                    oCmdEjec.CommandText = "SPPROGRAMACION_DETALLE_REGISTRAR"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , vIDP)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lvDetalle.ListItems(ip).Tag)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRODUCTO", adVarChar, adParamInput, 200, Me.lvDetalle.ListItems(ip).Text)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adBigInt, adParamInput, , Me.lvDetalle.ListItems(ip).SubItems(1))
                    oCmdEjec.Execute
                Next
              
            Else
                oCmdEjec.CommandText = "SPPROGRAMACION_MODIFICAR"
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDTURNO", adInteger, adParamInput, , Me.DatTurno.BoundText)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , Me.dtpFecha.Value)
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , xIDPROGRAMACION)
                oCmdEjec.Execute
                
                'ELIMINAR PROGRAMACION
                LimpiaParametros oCmdEjec
                oCmdEjec.CommandText = "SPPROGRAMACION_ELIMINAR"
                
                Dim i As Integer

                For i = 0 To Xim - 1
                    LimpiaParametros oCmdEjec
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , xIDPROGRAMACION)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , xIDp(i))
                    oCmdEjec.Execute
                
                Next
                
                For i = 1 To Me.lvDetalle.ListItems.count
                    LimpiaParametros oCmdEjec
                
                    oCmdEjec.CommandText = "SPPROGRAMACION_DETALLE_MODIFICAR"
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , xIDPROGRAMACION)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.lvDetalle.ListItems(i).Tag)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PRODUCTO", adVarChar, adParamInput, 200, Me.lvDetalle.ListItems(i).Text)
                    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adBigInt, adParamInput, , Me.lvDetalle.ListItems(i).SubItems(1))
                    oCmdEjec.Execute
                Next
                
            End If

            MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
            Pub_ConnAdo.CommitTrans
            Unload Me

            Exit Sub

Graba:
            Pub_ConnAdo.RollbackTrans
            MsgBox Err.Description + vbCrLf + Err.Source, vbCritical, Pub_Titulo

        Case 2 'CANCELAR
            Unload Me
    End Select
    
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdAgregar_Click
End Sub

Private Sub txtProducto_Change()
    vBuscar = True
End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
   On Error GoTo SALE

    Dim strFindMe As String

    Dim itmFound  As Object ' ListItem    ' Variable FoundItem.

    If KeyCode = 40 Then  ' flecha abajo
        loc_key = loc_key + 1

        If loc_key > Me.lvProductos.ListItems.count Then loc_key = lvProductos.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 38 Then
        loc_key = loc_key - 1

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 34 Then
        loc_key = loc_key + 17

        If loc_key > lvProductos.ListItems.count Then loc_key = lvProductos.ListItems.count
        GoTo POSICION
    End If

    If KeyCode = 33 Then
        loc_key = loc_key - 17

        If loc_key < 1 Then loc_key = 1
        GoTo POSICION
    End If

    If KeyCode = 27 Then
        Me.lvProductos.Visible = False
        Me.txtProducto.Text = ""
        '        Me.lblDocumento.Caption = ""
        '        Me.lblTelefonos.Caption = ""
    End If

    GoTo fin
POSICION:
    lvProductos.ListItems.Item(loc_key).Selected = True
    lvProductos.ListItems.Item(loc_key).EnsureVisible
    'txtRS.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
    Me.txtProducto.SelStart = Len(Me.txtProducto.Text)
fin:

    Exit Sub

SALE:
End Sub

Private Sub txtProducto_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        If vBuscar Then
            Me.lvProductos.ListItems.Clear
            LimpiaParametros oCmdEjec
            Dim orsDATA As ADODB.Recordset
            oCmdEjec.CommandText = "SPPROGRAMACION_PRODUCTOLIST"
            Set orsDATA = oCmdEjec.Execute(, Array(LK_CODCIA, Me.txtProducto.Text))

            Dim Item As Object
        
            If Not orsDATA.EOF Then

                Do While Not orsDATA.EOF
                    Set Item = Me.lvProductos.ListItems.Add(, , Trim(orsDATA!producto))
                    Item.Tag = Trim(orsDATA!Codigo)
                    
                    orsDATA.MoveNext
                Loop

                Me.lvProductos.Visible = True
                Me.lvProductos.ListItems(1).Selected = True
                loc_key = 1
                Me.lvProductos.ListItems(1).EnsureVisible
                vBuscar = False
            
                '         If MsgBox("Cliente no existe." + vbCrLf + "¿Desea Crearlo.?", vbQuestion + vbYesNo, "Restaurantes") = vbYes Then
                '         frmCLI.Show vbModal
                '         End If
            End If

        Else 'ASIGNAR VALORES A LOS CONTROLES

            If Me.txtProducto.Text <> "" Then
                Me.lvProductos.Visible = False
                Me.txtProducto.Text = Me.lvProductos.SelectedItem.Text
                Me.txtProducto.Tag = Me.lvProductos.SelectedItem.Tag
                'Me.txtPrecioPlato.Text = Me.lvProductos.SelectedItem.SubItems(1)
                'Me.FraDetalle.Enabled = True
                Me.txtCantidad.SetFocus
                Me.txtCantidad.SelStart = 0
                Me.txtCantidad.SelLength = Len(Me.txtCantidad.Text)
                
            
            End If
            
        End If
    End If
End Sub
