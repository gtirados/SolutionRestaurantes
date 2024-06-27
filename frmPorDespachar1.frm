VERSION 5.00
Object = "{BB35AEF3-E525-4F8B-81F2-511FF805ABB1}#2.1#0"; "scrollerii.ocx"
Begin VB.Form frmPorDespachar1 
   Caption         =   "Platos por Despachar"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13020
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrPedidos 
      Interval        =   5000
      Left            =   5880
      Top             =   3840
   End
   Begin VB.Frame FraPedido 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   6735
      Begin VB.CheckBox chkSegundos 
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   18
         Tag             =   "X"
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox chkEntradas 
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   17
         Tag             =   "X"
         Top             =   1800
         Width           =   255
      End
      Begin VB.CheckBox chkProducto 
         Caption         =   "Productosadsdasdasdasdasdasdasdasdddddd"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   3
         Top             =   2040
         Width           =   4815
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Tag             =   "X"
         Top             =   1770
         Width           =   1215
      End
      Begin VB.CommandButton cmdDespachar 
         Caption         =   "Despachar"
         Height          =   360
         Index           =   0
         Left            =   4980
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   22
         Top             =   960
         Width           =   4995
      End
      Begin VB.Label lbleDireccion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCIÓN:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label lblAdicional 
         BackStyle       =   0  'Transparent
         Caption         =   "989"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   20
         Top             =   2310
         Width           =   5835
      End
      Begin VB.Label lblTipo 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   6960
         TabIndex        =   19
         Top             =   1710
         Width           =   555
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "COMANDA NRO 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   240
         Width           =   6615
      End
      Begin VB.Label lbleMozo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOZO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   885
         TabIndex        =   15
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblMozo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   600
         Width           =   4995
      End
      Begin VB.Label lbleHoraInicio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HORA INICIO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1350
         Width           =   1485
      End
      Begin VB.Label lblHoraInicio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   1320
         Width           =   2115
      End
      Begin VB.Label lbleTiempo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIEMPO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4080
         TabIndex        =   11
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblTiempo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   5040
         TabIndex        =   10
         Top             =   1320
         Width           =   1635
      End
      Begin VB.Line lnLinea 
         Index           =   0
         X1              =   120
         X2              =   6720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblEntradas 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COCINA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   1755
         Width           =   6615
      End
      Begin VB.Label lblCantidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   7
         Top             =   2040
         Width           =   315
      End
      Begin VB.Label lblDetalle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle4556784ddddddd"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   5280
         TabIndex        =   6
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label lblSerie 
         BackStyle       =   0  'Transparent
         Height          =   315
         Index           =   0
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   930
      End
      Begin VB.Label lblNumero 
         BackStyle       =   0  'Transparent
         Height          =   315
         Index           =   0
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label lblSegundos 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   3600
         Width           =   6615
      End
   End
   Begin ScrollerII.FormScroller FormScroller1 
      Left            =   6120
      Top             =   7080
      _ExtentX        =   2170
      _ExtentY        =   1085
      SmallChange     =   100
      LargeChange     =   720
      BackColor       =   -2147483632
      ScaleMode       =   0
   End
End
Attribute VB_Name = "frmPorDespachar1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private vCarga As Boolean
Private vVERDELIVERY As Boolean
Private vTIPO As Integer
Private vCant As Integer
Private oRStemp As New ADODB.Recordset 'para los platos

Private Sub chkEntradas_Click(Index As Integer)

    Dim ElControl As Control
     
    'recorre los controles
      
    For Each ElControl In Controls

        'si está dentro lo deshabilita
        If ElControl.Name <> "FormScroller1" And ElControl.Name <> "tmrPedidos" Then
     
            If ElControl.Container Is FraPedido(Index) Then
               
                If TypeOf ElControl Is CheckBox And ElControl.Tag <> "X" And Me.lblTipo(ElControl.Index).Caption = "E" Then
                    If chkEntradas(Index).Value Then
                        ElControl.Value = 1
                    Else
                        ElControl.Value = 0
                    End If
                End If

                'ElControl.Enabled = False
            End If
        End If

    Next

End Sub

Private Sub chkProducto_Click(Index As Integer)

    'AQUI
    If Not vCarga Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPPENDIENTES_MARCA"

        Dim MyArray() As String

        MyArray = Split(chkProducto(Index).Tag, "|")
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.lblSerie(MyArray(1)).Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adInteger, adParamInput, , Me.lblNumero(MyArray(1)).Caption)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , MyArray(2))
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MARCA", adBoolean, adParamInput, , IIf(Me.chkProducto(Index).Value, 1, 0))
        oCmdEjec.Execute
    End If

End Sub

Private Sub chkSegundos_Click(Index As Integer)

    Dim ElControl As Control
     
    'recorre los controles
      
    For Each ElControl In Controls

        'si está dentro lo deshabilita
        If ElControl.Name <> "FormScroller1" And ElControl.Name <> "tmrPedidos" Then
     
            If ElControl.Container Is FraPedido(Index) Then
               
                If TypeOf ElControl Is CheckBox And ElControl.Tag <> "X" And Me.lblTipo(ElControl.Index).Caption = "S" Then
                    If chkSegundos(Index).Value Then
                        ElControl.Value = 1
                    Else
                        ElControl.Value = 0
                    End If
                End If

                'ElControl.Enabled = False
            End If
        End If

    Next
End Sub

Private Sub chkTodos_Click(Index As Integer)

    'Variable de tipo Control Para los controles del contenedor en este caso del Frame
    Dim ElControl As Control
     
    'recorre los controles
      
    For Each ElControl In Controls

        'si está dentro lo deshabilita
        If ElControl.Name <> "FormScroller1" And ElControl.Name <> "tmrPedidos" Then
     
            If ElControl.Container Is FraPedido(Index) Then
                If TypeOf ElControl Is CheckBox And ElControl.Tag <> "X" Then
                    If chkTodos(Index).Value Then
                        ElControl.Value = 1
                    Else
                        ElControl.Value = 0
                    End If
                End If
            End If
        End If

    Next

End Sub

Private Sub cmdDespachar_Click(Index As Integer)
Me.tmrPedidos.Enabled = False
Dim oRScodigos As New ADODB.Recordset
    
    oRScodigos.Fields.Append "CODIGO", adBigInt

    oRScodigos.CursorLocation = adUseClient
    oRScodigos.LockType = adLockOptimistic
    oRScodigos.CursorType = adOpenDynamic
    oRScodigos.Open
    

    Dim ElControl As Control
     
    'recorre los controles
      
    For Each ElControl In Controls

        'si está dentro lo deshabilita
        If ElControl.Name <> "FormScroller1" And ElControl.Name <> "tmrPedidos" Then
     
            If ElControl.Container Is FraPedido(Index) Then
                If TypeOf ElControl Is CheckBox And ElControl.Tag <> "X" Then
                    If ElControl.Value And ElControl.Enabled Then
                     Dim MyArray() As String
                     MyArray = Split(Me.chkProducto(ElControl.Index).Tag, "|")
                     
                     oRScodigos.Filter = "codigo=" & MyArray(0)
                     If oRScodigos.RecordCount = 0 Then
                     
                        LimpiaParametros oCmdEjec
                        oCmdEjec.CommandText = "SPCOMANDADESPACHAR"
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
                        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.lblSerie(Index).Caption)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lblNumero(Index).Caption)
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , Me.lblCantidad(ElControl.Index).Tag)
                       
                        
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODART", adBigInt, adParamInput, , MyArray(0))
                        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, , Me.lblCantidad(ElControl.Index).Caption)
                        oCmdEjec.Execute
                        oRScodigos.AddNew
                        oRScodigos!Codigo = MyArray(0)
                        oRScodigos.Update
                        
                        End If
                       
                    End If
                     vCant = vCant - 1
                End If

                'ElControl.Enabled = False
            End If
        End If

    Next
    
    CargarPedidos vVERDELIVERY, vTIPO
    Me.tmrPedidos.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
    If KeyCode = vbKeyF3 Then 'MUESTRA LOS PEDIDOS DELIVERY
        vCant = 0
        vVERDELIVERY = False
        CleanControles
        vTIPO = 3
        CargarPedidos False, vTIPO
    ElseIf KeyCode = vbKeyF4 Then 'MUESTRA LOS PEDIDOS NO DELIVERY
     
        
        vVERDELIVERY = True
        CleanControles
        CargarPedidos True
    ElseIf KeyCode = vbKeyF5 Then
        Me.tmrPedidos.Enabled = Not Me.tmrPedidos.Enabled

    ElseIf KeyCode = vbKeyF1 Then 'MUESTRA LOS PEDIDOS NO DELIVERY - cocina
      vCant = 0
        vVERDELIVERY = False
        CleanControles
        vTIPO = 1
        CargarPedidos False, vTIPO
    ElseIf KeyCode = vbKeyF2 Then 'MUESTRA LOS PEDIDOS NO DELIVERY - bar
        vCant = 0
        vVERDELIVERY = False
        CleanControles
        vTIPO = 2
        CargarPedidos False, vTIPO
    End If

End Sub

Private Sub CleanControles()

    Dim i As Integer

    For i = 1 To Me.lblTitulo.count - 1
        Unload Me.lblTitulo.Item(i)
        Unload Me.lbleMozo.Item(i)
        Unload Me.lbleHoraInicio.Item(i)
        Unload Me.lbleTiempo.Item(i)
        Unload Me.lblMozo.Item(i)
        Unload Me.lblHoraInicio.Item(i)
        Unload Me.lblTiempo.Item(i)
        Unload Me.lnLinea.Item(i)
        Unload Me.cmdDespachar.Item(i)
        Unload Me.chkTodos.Item(i)
        Unload Me.lblSerie.Item(i)
        Unload Me.lblNumero.Item(i)
    Next

    For i = 1 To Me.lblEntradas.count - 1
        Unload Me.lblEntradas.Item(i)
    Next
    
    '    For i = 1 To Me.chkEntradas.count - 1
    '        Unload Me.chkEntradas.Item(i)
    '    Next
    
    Dim x As Control

    For Each x In Controls

        If x.Name = "chkEntradas" Then
            If x.Index <> 0 Then
                Unload x
            End If
        End If

    Next

    For Each x In Controls

        If x.Name = "chkSegundos" Then
            If x.Index <> 0 Then
                Unload x
            End If
        End If

    Next

    For i = 1 To Me.lblSegundos.count - 1
        Unload Me.lblSegundos.Item(i)
    Next

    For i = 1 To Me.chkProducto.count - 1
        Unload Me.chkProducto.Item(i)
        Unload Me.lblCantidad.Item(i)
        Unload Me.lblDetalle.Item(i)
       
    Next

    For i = 1 To Me.lblTipo.count - 1
        Unload Me.lblTipo.Item(i)
    Next

    For i = 1 To Me.lblDireccion.count - 1
        Unload Me.lblDireccion.Item(i)
    Next

    For i = 1 To Me.lbleDireccion.count - 1
        Unload Me.lbleDireccion.Item(i)
    Next

    For i = 1 To Me.lblAdicional.count - 1
        Unload Me.lblAdicional.Item(i)
    Next

    For i = 1 To Me.FraPedido.count - 1
        Unload Me.FraPedido.Item(i)
    Next

End Sub

Private Sub Form_Load()

Set oRStemp = New ADODB.Recordset

oRStemp.Fields.Append "PLATO", adVarChar, 100

oRStemp.CursorLocation = adUseClient
oRStemp.LockType = adLockOptimistic
oRStemp.CursorType = adOpenDynamic
oRStemp.Open
    
    
vCant = 0
vCarga = True
CargarPedidos False
vCarga = False
vTIPO = 3

  
    
End Sub

Private Sub Form_Resize()
    'Me.ScrollBox1.Width = Me.ScaleWidth
    'Me.ScrollBox1.Height = Me.ScaleHeight
End Sub

Private Sub CargarPedidos(xDelivery As Boolean, Optional xtipo As Integer = 3)
   
    Dim i    As Integer

    Dim ca   As Integer

    Dim xAdi As Integer

    xAdi = 0
    CleanControles
  
    LimpiaParametros oCmdEjec

    Dim orsPendientes As ADODB.Recordset

    Dim oRsPlatos     As ADODB.Recordset

    If xDelivery Then
        oCmdEjec.CommandText = "SPITEMSxDESPACHAR_DELIVERY"
    Else
        oCmdEjec.CommandText = "SPITEMSxDESPACHAR1"
    End If

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Fecha", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)

    If Not xDelivery Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@Tipo", adTinyInt, adParamInput, , xtipo)
    End If

    Set orsPendientes = oCmdEjec.Execute
    Set oRsPlatos = orsPendientes.NextRecordset

    Dim xSuena As Boolean

    xSuena = False

    If oRStemp.RecordCount = 0 And oRsPlatos.RecordCount <> 0 Then
        xSuena = True

        Do While Not oRsPlatos.EOF
            
            oRStemp.AddNew
            oRStemp!plato = oRsPlatos!producto
            oRStemp.Update
        
            oRsPlatos.MoveNext
        Loop

    Else

        Do While Not oRsPlatos.EOF
        
            oRStemp.Filter = "PLATO='" & oRsPlatos!producto & "'"
        
            If oRStemp.RecordCount = 0 Then
                xSuena = True

                Exit Do

            End If

            oRStemp.Filter = ""

            oRsPlatos.MoveNext
        Loop

    End If
    
    If xSuena Then
        PlaySound "C:\Admin\Sonidos\alarm.wav.", 1, 1
    End If
    
    Do While Not oRStemp.EOF
        oRStemp.Delete adAffectCurrent
        oRStemp.MoveNext
    Loop
            
    If oRsPlatos.RecordCount <> 0 Then oRsPlatos.MoveFirst

    Do While Not oRsPlatos.EOF
            
        oRStemp.AddNew
        oRStemp!plato = oRsPlatos!producto
        oRStemp.Update
        
        oRsPlatos.MoveNext
    Loop
    
    '    If oRsPlatos.RecordCount > vCant Then
    '        vCant = oRsPlatos.RecordCount
    '        PlaySound "C:\Admin\Sonidos\alarm.wav.", 1, 1
    '    End If
    
    Dim vColumna As Integer 'cantidad de columnas que se mostraran en pantalla

    Dim vLeft, vTop1, vTop2, vTop3, vTopPlato As Integer

    Dim EE, eS, CP As Integer

    Dim AltoFrame As Integer

    AltoFrame = 2175

    vLeft = 0
    vTop1 = 0
    vTop2 = 0
    vTop3 = 0
    vTopPlato = 0
    i = 1
    EE = 1
    eS = 1
    CP = 1
    ca = 1
    vColumna = 1

    'ADICIONAL
    Dim xfile      As Integer

    Dim xe         As String

    Dim xarraye()  As String

    Dim xheighadie As Double

    If xDelivery Then

        Do While Not orsPendientes.EOF
            Load FraPedido(i)

            If i > 3 Then 'a partir de la 2 fila
                If vColumna = 1 Then vTop1 = vTop1 + FraPedido(i - 3).Height
                If vColumna = 2 Then vTop2 = vTop2 + FraPedido(i - 3).Height
                If vColumna = 3 Then vTop3 = vTop3 + FraPedido(i - 3).Height
            End If
            
            If vColumna = 1 Then
                vLeft = 0
            Else
                vLeft = vLeft + FraPedido(i - 1).Width
            
            End If

            FraPedido(i).Left = vLeft

            If vColumna = 1 Then
                FraPedido(i).Top = vTop1
            ElseIf vColumna = 2 Then
                FraPedido(i).Top = vTop2
            ElseIf vColumna = 3 Then
                FraPedido(i).Top = vTop3
            End If
        
            FraPedido(i).Visible = True
        
            'AGREGANDO EL TITULO
            Load lblTitulo(i)
            Set lblTitulo(i).Container = FraPedido(i)
            lblTitulo(i).Caption = CStr(orsPendientes!serie) + "-" + CStr(orsPendientes!NUMERO)
            lblTitulo(i).Visible = True
            lblTitulo(i).Top = Me.lblTitulo(0).Top
            lblTitulo(i).Left = Me.lblTitulo(0).Left
            lblTitulo(i).Alignment = 2
            lblTitulo(i).Width = Me.lblTitulo(0).Width
            lblTitulo(i).AutoSize = False
            'AGREGANDO SERIE Y NRO DE COMANDA
            Load lblSerie(i)
            Set lblSerie(i).Container = FraPedido(i)

            With Me.lblSerie(i)
                .Caption = orsPendientes!serie
                .Width = Me.lblSerie(0).Width
                .Top = Me.lblSerie(0).Top
                .Height = Me.lblSerie(0).Height
                .Visible = True
            End With

            Load lblNumero(i)
            Set lblNumero(i).Container = FraPedido(i)

            With Me.lblNumero(i)
                .Caption = orsPendientes!NUMERO
                .Width = Me.lblNumero(0).Width
                .Top = Me.lblNumero(0).Top
                .Height = Me.lblNumero(0).Height
                .Visible = True
            End With
        
            'AGREGANDO ETIQUETA MOZO Y MOZO
            Load lbleMozo(i)
            Set lbleMozo(i).Container = FraPedido(i)

            With lbleMozo(i)
                .Caption = "CLIENTE:"
                .Visible = True
                .Top = Me.lbleMozo(0).Top
                .Left = Me.lbleMozo(0).Left
                .Width = Me.lbleMozo(0).Width
            End With

            Load lblMozo(i)
            Set lblMozo(i).Container = FraPedido(i)

            With lblMozo(i)
                .Caption = orsPendientes!mozo
                .Visible = True
                .Top = Me.lblMozo(0).Top
                .Left = Me.lblMozo(0).Left
                .Width = Me.lblMozo(0).Width
            End With
            
            'AGREGANDO ETIQUETA DIRECCION
            Load lbleDireccion(i)
            Set lbleDireccion(i).Container = FraPedido(i)

            With lbleDireccion(i)
                .Caption = "DIRECCION:"
                .Visible = True
                .Top = Me.lbleDireccion(0).Top
                .Left = Me.lbleDireccion(0).Left
                .Width = Me.lbleDireccion(0).Width
            End With

            Load lblDireccion(i)
            Set lblDireccion(i).Container = FraPedido(i)

            With lblDireccion(i)
                .Caption = orsPendientes!direccion
                .Visible = True
                .Top = Me.lblDireccion(0).Top
                .Left = Me.lblDireccion(0).Left
                .Width = Me.lblDireccion(0).Width
            End With

            'AGREGANDO ETIQUETA HORA INICIO
            Load lbleHoraInicio(i)
            Set lbleHoraInicio(i).Container = FraPedido(i)

            With lbleHoraInicio(i)
                .Caption = Me.lbleHoraInicio(0).Caption
                .Visible = True
                .Top = Me.lbleHoraInicio(0).Top
                .Left = Me.lbleHoraInicio(0).Left
                .Width = Me.lbleHoraInicio(0).Width
            End With

            Load lblHoraInicio(i)
            Set lblHoraInicio(i).Container = FraPedido(i)

            With lblHoraInicio(i)
                .Caption = orsPendientes!REGISTRO
                .Visible = True
                .Top = Me.lblHoraInicio(0).Top
                .Left = Me.lblHoraInicio(0).Left
                .Width = Me.lblHoraInicio(0).Width
            End With

            'AGREGANDO ETIQUETA DE TIEMPO
            Load lbleTiempo(i)
            Set lbleTiempo(i).Container = FraPedido(i)

            With lbleTiempo(i)
                .Caption = Me.lbleTiempo(0).Caption
                .Visible = True
                .Top = Me.lbleTiempo(0).Top
                .Left = Me.lbleTiempo(0).Left
                .Width = Me.lbleTiempo(0).Width
            End With

            Load lblTiempo(i)
            Set lblTiempo(i).Container = FraPedido(i)

            With lblTiempo(i)
                .Caption = orsPendientes!tiempo
                .Visible = True
                .Top = Me.lblTiempo(0).Top
                .Left = Me.lblTiempo(0).Left
                .Width = Me.lblTiempo(0).Width
            End With

            'AGREGANDO LA LINEA
            Load lnLinea(i)
            Set lnLinea(i).Container = FraPedido(i)

            With lnLinea(i)
                .Visible = True
                .X1 = Me.lnLinea(0).X1
                .X2 = Me.lnLinea(0).X2
                .Y1 = Me.lnLinea(0).Y1
                .Y2 = Me.lnLinea(0).Y2
            End With

            'AGREGANDO CHECK TODOS
            Load chkTodos(i)
            Set chkTodos(i).Container = FraPedido(i)

            With chkTodos(i)
                .Visible = True
                .Tag = Me.chkTodos(0).Tag
                .Caption = Me.chkTodos(0).Caption
                .Top = Me.chkTodos(0).Top
                .Left = Me.chkTodos(0).Left
                .Width = Me.chkTodos(0).Width
            End With

            'AGREGANDO BOTON PARA DEPACHAR
            Load cmdDespachar(i)
            Set cmdDespachar(i).Container = FraPedido(i)

            With cmdDespachar(i)
                .Visible = True
                .Caption = cmdDespachar(0).Caption
                .Top = cmdDespachar(0).Top
                .Left = cmdDespachar(0).Left
                .Width = cmdDespachar(0).Width
            End With
            
            'AGREGANDO ENTRADAS
            vTopPlato = Me.lblEntradas(0).Height + Me.lblEntradas(0).Top + 10

            oRsPlatos.Filter = "numero=" & CStr(orsPendientes!NUMERO) & " And SERIE=" & CStr(orsPendientes!serie)
                
            Do While Not oRsPlatos.EOF
                'AGREGA PLATO
                Load chkProducto(CP)
                Set chkProducto(CP).Container = FraPedido(i)

                With chkProducto(CP)
                    .Visible = True
                    .Caption = oRsPlatos!producto
                    .ToolTipText = oRsPlatos!productotool
                    .Tag = oRsPlatos!Codigo & "|" & i & "|" & oRsPlatos!SEC
                    .Top = vTopPlato
                    .Left = Me.chkProducto(0).Left
                    .Width = Me.chkProducto(0).Width
                    .Value = IIf(oRsPlatos!marca, 1, 0)
                    .ForeColor = &H8000&
                    .FontBold = True
                   ' .Enabled = oRsPlatos!Cc
                End With

                'AGREGANDO LA CANTIDAD
                Load lblCantidad(CP)
                Set Me.lblCantidad(CP).Container = FraPedido(i)

                With Me.lblCantidad(CP)
                    .Visible = True
                    .Caption = oRsPlatos!Cantidad
                    .Tag = oRsPlatos!SEC
                    .Top = vTopPlato
                    .Left = Me.lblCantidad(0).Left
                    .Width = Me.lblCantidad(0).Width
                    .ForeColor = &H8000&
                    .FontBold = True
                End With

                'AGREGANDO DETALLE
                Load lblDetalle(CP)
                Set Me.lblDetalle(CP).Container = FraPedido(i)

                With Me.lblDetalle(CP)
                    .Visible = True
                    .Caption = oRsPlatos!DETALLE
                    .Top = vTopPlato
                    .Left = Me.lblDetalle(0).Left
                    .Width = Me.lblDetalle(0).Width
                    .ForeColor = &H8000&
                    .FontBold = True
                End With

                'AGREGANDO EL TIPO POR CADA PRODUCTO
                Load lblTipo(CP)
                Set Me.lblTipo(CP).Container = FraPedido(i)

                With Me.lblTipo(CP)
                    '.Visible = True
                    .Caption = "E"
                    .Top = vTopPlato
                    .Left = Me.lblTipo(0).Left
                    .Width = Me.lblTipo(0).Width
                End With
          
                xfile = 0
                xheighadie = 0
                    
                If Len(Trim(oRsPlatos!ADICIONAL)) <> 0 Then
                    vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10
                    Load Me.lblAdicional(ca)
                    Set Me.lblAdicional(ca).Container = FraPedido(i)

                    xe = Len(Trim(oRsPlatos!ADICIONAL)) / 55
                    xarraye = Split(xe, ".")

                    If UBound(xarraye) <> 0 Then
                        If xarraye(1) <> 0 Then
                            xfile = xfile + xarraye(0) + 1
                        End If
                    End If

                    xAdi = xAdi + xfile

                    With Me.lblAdicional(ca)
                        .Caption = oRsPlatos!ADICIONAL
                        .Left = Me.lblAdicional(0).Left
                        .Width = Me.lblAdicional(0).Width
                        .Height = xfile * 195
                        
                        xheighadie = .Height
                        .Top = vTopPlato
                        .Visible = True
                        ' MsgBox "ca"
                      
                    End With

                    ca = ca + 1
                End If

                vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10 + IIf(xheighadie <> 0, xheighadie - 195, 0)
                oRsPlatos.MoveNext
                CP = CP + 1
            Loop
            
            vColumna = vColumna + 1
            vTopPlato = 0

            If orsPendientes!PLATOS <> 0 Then
                AltoFrame = AltoFrame + ((orsPendientes!PLATOS) * 255)
            End If
        
            If xAdi <> 0 Then
                AltoFrame = AltoFrame + (xAdi * 200)
            End If

            FraPedido(i).Height = AltoFrame

            If vColumna > 3 Then vColumna = 1
            i = i + 1
            orsPendientes.MoveNext
            AltoFrame = 2175
            
        Loop

    Else

        Do While Not orsPendientes.EOF
            Load FraPedido(i)
        
            If i > 3 Then 'a partir de la 2 fila
                If vColumna = 1 Then vTop1 = vTop1 + FraPedido(i - 3).Height
                If vColumna = 2 Then vTop2 = vTop2 + FraPedido(i - 3).Height
                If vColumna = 3 Then vTop3 = vTop3 + FraPedido(i - 3).Height
            End If

            If vColumna = 1 Then
                vLeft = 0
            Else
                vLeft = vLeft + FraPedido(i - 1).Width
            
            End If

            FraPedido(i).Left = vLeft

            If vColumna = 1 Then
                FraPedido(i).Top = vTop1
            ElseIf vColumna = 2 Then
                FraPedido(i).Top = vTop2
            ElseIf vColumna = 3 Then
                FraPedido(i).Top = vTop3
            End If
        
            FraPedido(i).Visible = True
        
            'AGREGANDO EL TITULO
            Load lblTitulo(i)
            Set lblTitulo(i).Container = FraPedido(i)
            lblTitulo(i).Caption = orsPendientes!mesa
            lblTitulo(i).Visible = True
            lblTitulo(i).Top = Me.lblTitulo(0).Top
            lblTitulo(i).Left = Me.lblTitulo(0).Left
            lblTitulo(i).Alignment = 2
            lblTitulo(i).Width = Me.lblTitulo(0).Width
            lblTitulo(i).AutoSize = False
            'AGREGANDO SERIE Y NRO DE COMANDA
            Load lblSerie(i)
            Set lblSerie(i).Container = FraPedido(i)

            With Me.lblSerie(i)
                .Caption = orsPendientes!serie
                .Width = Me.lblSerie(0).Width
                .Top = Me.lblSerie(0).Top
                .Height = Me.lblSerie(0).Height
                .Visible = True
            End With

            Load lblNumero(i)
            Set lblNumero(i).Container = FraPedido(i)

            With Me.lblNumero(i)
                .Caption = orsPendientes!NUMERO
                .Width = Me.lblNumero(0).Width
                .Top = Me.lblNumero(0).Top
                .Height = Me.lblNumero(0).Height
                .Visible = True
            End With
        
            'AGREGANDO ETIQUETA MOZO Y MOZO
            Load lbleMozo(i)
            Set lbleMozo(i).Container = FraPedido(i)

            With lbleMozo(i)
                .Caption = "MOZO:"
                .Visible = True
                .Top = Me.lbleMozo(0).Top
                .Left = Me.lbleMozo(0).Left
                .Width = Me.lbleMozo(0).Width
            End With

            Load lblMozo(i)
            Set lblMozo(i).Container = FraPedido(i)

            With lblMozo(i)
                .Caption = orsPendientes!mozo
                .Visible = True
                .Top = Me.lblMozo(0).Top
                .Left = Me.lblMozo(0).Left
                .Width = Me.lblMozo(0).Width
            End With

            'AGREGANDO ETIQUETA HORA INICIO
            Load lbleHoraInicio(i)
            Set lbleHoraInicio(i).Container = FraPedido(i)

            With lbleHoraInicio(i)
                .Caption = Me.lbleHoraInicio(0).Caption
                .Visible = True
                .Top = Me.lbleHoraInicio(0).Top
                .Left = Me.lbleHoraInicio(0).Left
                .Width = Me.lbleHoraInicio(0).Width
            End With

            Load lblHoraInicio(i)
            Set lblHoraInicio(i).Container = FraPedido(i)

            With lblHoraInicio(i)
                .Caption = orsPendientes!REGISTRO
                .Visible = True
                .Top = Me.lblHoraInicio(0).Top
                .Left = Me.lblHoraInicio(0).Left
                .Width = Me.lblHoraInicio(0).Width
            End With

            'AGREGANDO ETIQUETA DE TIEMPO
            Load lbleTiempo(i)
            Set lbleTiempo(i).Container = FraPedido(i)

            With lbleTiempo(i)
                .Caption = Me.lbleTiempo(0).Caption
                .Visible = True
                .Top = Me.lbleTiempo(0).Top
                .Left = Me.lbleTiempo(0).Left
                .Width = Me.lbleTiempo(0).Width
            End With

            Load lblTiempo(i)
            Set lblTiempo(i).Container = FraPedido(i)

            With lblTiempo(i)
                .Caption = orsPendientes!tiempo
                .Visible = True
                .Top = Me.lblTiempo(0).Top
                .Left = Me.lblTiempo(0).Left
                .Width = Me.lblTiempo(0).Width
            End With

            'AGREGANDO LA LINEA
            Load lnLinea(i)
            Set lnLinea(i).Container = FraPedido(i)

            With lnLinea(i)
                .Visible = True
                .X1 = Me.lnLinea(0).X1
                .X2 = Me.lnLinea(0).X2
                .Y1 = Me.lnLinea(0).Y1
                .Y2 = Me.lnLinea(0).Y2
            End With

            'AGREGANDO CHECK TODOS
            Load chkTodos(i)
            Set chkTodos(i).Container = FraPedido(i)

            With chkTodos(i)
                .Visible = True
                .Tag = Me.chkTodos(0).Tag
                .Caption = Me.chkTodos(0).Caption
                .Top = Me.chkTodos(0).Top
                .Left = Me.chkTodos(0).Left
                .Width = Me.chkTodos(0).Width
            End With

            'AGREGANDO BOTON PARA DEPACHAR
            Load cmdDespachar(i)
            Set cmdDespachar(i).Container = FraPedido(i)

            With cmdDespachar(i)
                .Visible = True
                .Caption = cmdDespachar(0).Caption
                .Top = cmdDespachar(0).Top
                .Left = cmdDespachar(0).Left
                .Width = cmdDespachar(0).Width
            End With

            'VERIFICANDO SI LA COMANDA TIENE ENTRADAS Y SEGUNDOS
            If orsPendientes!ENTRADAS <> 0 Then
                'AGREGANDO TITULO PARA ENTRADAS
                Load lblEntradas(EE)
                Set lblEntradas(EE).Container = FraPedido(i)

                With lblEntradas(EE)
                    .Visible = True
                    .Caption = "COCINA"
                    .Top = Me.lblEntradas(0).Top
                    .Left = Me.lblEntradas(0).Left
                    .Width = Me.lblEntradas(0).Width
                End With

                'AGREGANDO EL CHECK PARA LAS ENTRADAS
                Load chkEntradas(i)
                Set Me.chkEntradas(i).Container = FraPedido(i)

                With Me.chkEntradas(i)
                    .Visible = True
                    .Tag = Me.chkEntradas(0).Tag
                    .Top = Me.chkEntradas(0).Top
                    .Left = Me.chkEntradas(0).Left
                    .Width = Me.chkEntradas(0).Width
                
                End With

                'AGREGANDO ENTRADAS
                oRsPlatos.Filter = "numero=" & CStr(orsPendientes!NUMERO) & " And FAMILIA=1"
                vTopPlato = Me.lblEntradas(0).Height + Me.lblEntradas(0).Top + 10
                Me.lblEntradas(EE).Caption = oRsPlatos!NOMFamilia

                Do While Not oRsPlatos.EOF
                    'AGREGA PLATO
                    Load chkProducto(CP)
                    Set chkProducto(CP).Container = FraPedido(i)

                    With chkProducto(CP)
                        .Visible = True
                        .Caption = oRsPlatos!producto
                        .ToolTipText = oRsPlatos!productotool
                        .Tag = oRsPlatos!Codigo & "|" & i & "|" & oRsPlatos!SEC
                        .Top = vTopPlato
                        .Left = Me.chkProducto(0).Left
                        .Width = Me.chkProducto(0).Width
                        .Value = IIf(oRsPlatos!marca, 1, 0)
                        .ForeColor = &H8000&
                        .FontBold = True
                        .Enabled = oRsPlatos!Cc
                    End With

                    'AGREGANDO LA CANTIDAD
                    Load lblCantidad(CP)
                    Set Me.lblCantidad(CP).Container = FraPedido(i)

                    With Me.lblCantidad(CP)
                        .Visible = True
                        .Caption = oRsPlatos!Cantidad
                        .Tag = oRsPlatos!SEC
                        .Top = vTopPlato
                        .Left = Me.lblCantidad(0).Left
                        .Width = Me.lblCantidad(0).Width
                        .ForeColor = &H8000&
                        .FontBold = True
                    End With

                    'AGREGANDO DETALLE
                    Load lblDetalle(CP)
                    Set Me.lblDetalle(CP).Container = FraPedido(i)

                    With Me.lblDetalle(CP)
                        .Visible = True
                        .Caption = oRsPlatos!DETALLE
                        .Top = vTopPlato
                        .Left = Me.lblDetalle(0).Left
                        .Width = Me.lblDetalle(0).Width
                        .ForeColor = &H8000&
                        .FontBold = True
                    End With

                    'AGREGANDO EL TIPO POR CADA PRODUCTO
                    Load lblTipo(CP)
                    Set Me.lblTipo(CP).Container = FraPedido(i)

                    With Me.lblTipo(CP)
                        '.Visible = True
                        .Caption = "E"
                        .Top = vTopPlato
                        .Left = Me.lblTipo(0).Left
                        .Width = Me.lblTipo(0).Width
                    End With

                    'ADICIONAL
                    '                    Dim xfile As Integer
                    '
                    '                    Dim xe    As String
                    '
                    '
                    '                    Dim xarraye()  As String
                    '
                    '                    Dim xheighadie As Double
                    xfile = 0

                    xheighadie = 0
                    
                    If Len(Trim(oRsPlatos!ADICIONAL)) <> 0 Then
                        vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10
                        Load Me.lblAdicional(ca)
                        Set Me.lblAdicional(ca).Container = FraPedido(i)

                        xe = Len(Trim(oRsPlatos!ADICIONAL)) / 55
                        xarraye = Split(xe, ".")

                        If UBound(xarraye) <> 0 Then
                            If xarraye(1) <> 0 Then
                                xfile = xfile + xarraye(0) + 1
                            End If
                        End If

                        xAdi = xAdi + xfile

                        With Me.lblAdicional(ca)
                            .Caption = oRsPlatos!ADICIONAL
                            .Left = Me.lblAdicional(0).Left
                            .Width = Me.lblAdicional(0).Width
                            .Height = xfile * 195
                        
                            xheighadie = .Height
                            .Top = vTopPlato
                            .Visible = True
                            ' MsgBox "ca"
                      
                        End With

                        ca = ca + 1
                    End If

                    vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10 + IIf(xheighadie <> 0, xheighadie - 195, 0)
                    oRsPlatos.MoveNext
                    CP = CP + 1
                Loop

                EE = EE + 1
                'AGREGANDO LOS SEGUNDOS
          
                oRsPlatos.Filter = "numero=" & CStr(orsPendientes!NUMERO) & " and familia=2"

                If oRsPlatos.RecordCount <> 0 Then
                    Load lblSegundos(eS)
                    Set lblSegundos(eS).Container = FraPedido(i)

                    With lblSegundos(eS)
                        .Visible = True
                        .Caption = oRsPlatos!NOMFamilia '"BAR"
                        .Top = vTopPlato
                        .Left = Me.lblEntradas(0).Left
                        .Width = Me.lblEntradas(0).Width
                    End With

                    'AGREGANDO EL CHECK PARA LOS SEGUNDOS
                    Load chkSegundos(i)
                    Set Me.chkSegundos(i).Container = FraPedido(i)

                    With Me.chkSegundos(i)
                        .Visible = True
                        '.Tag = "S"
                        .Tag = Me.chkSegundos(0).Tag
                        .Top = vTopPlato
                        .Value = IIf(oRsPlatos!marca, 1, 0)
                        .Left = Me.chkSegundos(0).Left
                        .Width = Me.chkSegundos(0).Width
                    End With

                    vTopPlato = vTopPlato + Me.lblEntradas(0).Height + 10

                    Do While Not oRsPlatos.EOF
                        'AGREGA PLATO Y CODIGOEN EL TAG
                        Load chkProducto(CP)
                        Set chkProducto(CP).Container = FraPedido(i)

                        With chkProducto(CP)
                            .Visible = True
                            .Caption = oRsPlatos!producto
                            .ToolTipText = oRsPlatos!productotool
                            .Tag = oRsPlatos!Codigo & "|" & i & "|" & oRsPlatos!SEC
                            .Top = vTopPlato
                            .Value = IIf(oRsPlatos!marca, 1, 0)
                            .Left = Me.chkProducto(0).Left
                            .Width = Me.chkProducto(0).Width
                            .ForeColor = vbRed
                            .FontBold = True
                            .Enabled = oRsPlatos!Cc
                        End With

                        'AGREGANDO LA CANTIDAD Y NRO DE SECUECIA EN EL TAG
                        Load lblCantidad(CP)
                        Set Me.lblCantidad(CP).Container = FraPedido(i)

                        With Me.lblCantidad(CP)
                            .Visible = True
                            .Caption = oRsPlatos!Cantidad
                            .Tag = oRsPlatos!SEC
                            .Top = vTopPlato
                            .Left = Me.lblCantidad(0).Left
                            .Width = Me.lblCantidad(0).Width
                            .ForeColor = vbRed
                            .FontBold = True
                        End With

                        'AGREGANDO DETALLE
                        Load lblDetalle(CP)
                        Set Me.lblDetalle(CP).Container = FraPedido(i)

                        With Me.lblDetalle(CP)
                            .Visible = True
                            .Caption = oRsPlatos!DETALLE
                            .Top = vTopPlato
                            .Left = Me.lblDetalle(0).Left
                            .Width = Me.lblDetalle(0).Width
                            .ForeColor = vbRed
                            .FontBold = True
                        End With

                        'AGREGANDO EL TIPO POR CADA PRODUCTO
                        Load lblTipo(CP)
                        Set Me.lblTipo(CP).Container = FraPedido(i)

                        With Me.lblTipo(CP)
                            '.Visible = True
                            .Caption = "S"
                            .Top = vTopPlato
                            .Left = Me.lblTipo(0).Left
                            .Width = Me.lblTipo(0).Width
                        End With
                    
                        'ADICIONAL
                        Dim xfils As Integer

                        Dim xs    As String

                        xfils = 0

                        Dim xarrays()  As String

                        Dim xheighadis As Double

                        xheighadis = 0
                    
                        If Len(Trim(oRsPlatos!ADICIONAL)) <> 0 Then
                            vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10
                            Load Me.lblAdicional(ca)
                            Set Me.lblAdicional(ca).Container = FraPedido(i)

                            xs = Len(Trim(oRsPlatos!ADICIONAL)) / 55
                            xarrays = Split(xs, ".")

                            If UBound(xarrays) <> 0 Then
                                If xarrays(1) <> 0 Then
                                    xfils = xfils + xarrays(0) + 1
                                End If
                            End If

                            xAdi = xAdi + xfils

                            With Me.lblAdicional(ca)
                                .Caption = oRsPlatos!ADICIONAL
                                .Left = Me.lblAdicional(0).Left
                                .Width = Me.lblAdicional(0).Width
                                .Height = xfils * 195
                        
                                xheighadis = .Height
                                .Top = vTopPlato
                                .Visible = True
                                ' MsgBox "ca"
                      
                            End With

                            ca = ca + 1
                        End If

                        vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10 + IIf(xheighadis <> 0, xheighadis - 195, 0)
                
                        oRsPlatos.MoveNext
                        CP = CP + 1
                    Loop

                    eS = eS + 1
                End If
            
            Else
                'AGREGANDO SOLO LOS SEGUNDOS
                Load lblSegundos(eS)
                Set Me.lblSegundos(eS).Container = FraPedido(i)

                With Me.lblSegundos(eS)
                    .Visible = True
                    .Caption = "BAR"
                    .Top = Me.lblEntradas(0).Top
                    .Left = Me.lblEntradas(0).Left
                    .Width = Me.lblEntradas(0).Width
                End With

                'AGREGANDO EL CHECK PARA LOS SEGUNDOS
                Load chkSegundos(i)
                Set Me.chkSegundos(i).Container = FraPedido(i)

                With Me.chkSegundos(i)
                    .Visible = True
                
                    .Tag = Me.chkSegundos(0).Tag
                    .Top = Me.lblEntradas(0).Top
                    .Left = Me.chkSegundos(0).Left
                    .Width = Me.chkSegundos(0).Width

                End With

                vTopPlato = chkProducto(0).Top + 10
                oRsPlatos.Filter = "numero=" & CStr(orsPendientes!NUMERO) & " and familia=2"
                Me.lblSegundos(eS).Caption = oRsPlatos!NOMFamilia

                Do While Not oRsPlatos.EOF
                    Load chkProducto(CP)
                    Set chkProducto(CP).Container = FraPedido(i)

                    With chkProducto(CP)
                        .Visible = True
                        .Caption = oRsPlatos!producto
                        .Tag = oRsPlatos!Codigo & "|" & i & "|" & oRsPlatos!SEC
                        .Value = IIf(oRsPlatos!marca, 1, 0)
                        .Top = vTopPlato
                        .Left = Me.chkProducto(0).Left
                        .Width = Me.chkProducto(0).Width
                        .ForeColor = vbRed
                        .FontBold = True
                        .Enabled = oRsPlatos!Cc
                    End With

                    'MsgBox "d"

                    'AGREGANDO LA CANTIDAD
                    Load lblCantidad(CP)
                    Set Me.lblCantidad(CP).Container = FraPedido(i)

                    With Me.lblCantidad(CP)
                        .Visible = True
                        .Caption = oRsPlatos!Cantidad
                        .Tag = oRsPlatos!SEC
                        .Top = vTopPlato
                        .Left = Me.lblCantidad(0).Left
                        .Width = Me.lblCantidad(0).Width
                        .ForeColor = vbRed
                        .FontBold = True
                    End With

                    'AGREGANDO DETALLE
                    Load lblDetalle(CP)
                    Set Me.lblDetalle(CP).Container = FraPedido(i)

                    With Me.lblDetalle(CP)
                        .Visible = True
                        .Caption = oRsPlatos!DETALLE
                        .Top = vTopPlato
                        .Left = Me.lblDetalle(0).Left
                        .Width = Me.lblDetalle(0).Width
                        .ForeColor = vbRed
                        .FontBold = True
                    End With

                    'AGREGANDO EL TIPO POR CADA PRODUCTO
                    Load lblTipo(CP)
                    Set Me.lblTipo(CP).Container = FraPedido(i)

                    With Me.lblTipo(CP)
                        '.Visible = True
                        .Caption = "S"
                        .Top = vTopPlato
                        .Left = Me.lblTipo(0).Left
                        .Width = Me.lblTipo(0).Width
                    End With
                
                    'ADICIONAL
                    Dim xfil As Integer

                    Dim x    As String

                    xfil = 0

                    Dim xarray()  As String

                    Dim xheighadi As Double

                    xheighadi = 0
                    
                    If Len(Trim(oRsPlatos!ADICIONAL)) <> 0 Then
                        vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10
                        Load Me.lblAdicional(ca)
                        Set Me.lblAdicional(ca).Container = FraPedido(i)

                        x = Len(Trim(oRsPlatos!ADICIONAL)) / 55
                        xarray = Split(x, ".")

                        If UBound(xarray) <> 0 Then
                            If xarray(1) <> 0 Then
                                xfil = xfil + xarray(0) + 1
                            End If
                        End If

                        xAdi = xAdi + xfil

                        With Me.lblAdicional(ca)
                            .Caption = oRsPlatos!ADICIONAL
                            .Left = Me.lblAdicional(0).Left
                            .Width = Me.lblAdicional(0).Width
                            .Height = xfil * 195
                        
                            xheighadi = .Height
                            .Top = vTopPlato
                            .Visible = True
                            ' MsgBox "ca"
                            ca = ca + 1
                        End With

                    End If

                    vTopPlato = vTopPlato + Me.chkProducto(CP).Height + 10 + xheighadi '- 195
                    'vTopPlato = vTopPlato + 10 + xheighadi
                    CP = CP + 1
                    oRsPlatos.MoveNext
                
                Loop

                eS = eS + 1
            End If

            vColumna = vColumna + 1
            vTopPlato = 0

            If orsPendientes!ENTRADAS <> 0 Then
                AltoFrame = AltoFrame + ((orsPendientes!ENTRADAS) * 320)
            End If

            If orsPendientes!SEGUNDOS <> 0 Then
                AltoFrame = AltoFrame + ((orsPendientes!SEGUNDOS) * 320)
            End If
        
            If xAdi <> 0 Then
                AltoFrame = AltoFrame + (xAdi * 200)
            End If

            FraPedido(i).Height = AltoFrame

            If vColumna > 3 Then vColumna = 1
            i = i + 1
            orsPendientes.MoveNext
            AltoFrame = 2175
        Loop

    End If

    'Me.FormScroller1.ContinuousScroll = False
    'Me.FormScroller1.ContinuousScroll = True
    '
    'Me.ScaleHeight = Me.ScaleHeight - 100
    '  Me.WindowState = 0
    ' Me.Height = Me.Height - 100
    ' Me.WindowState = 2
 
End Sub


Private Sub tmrPedidos_Timer()
vCarga = True
'FALTA AGREGAR EL PARAMETRO
CargarPedidos vVERDELIVERY, vTIPO
vCarga = False
End Sub
