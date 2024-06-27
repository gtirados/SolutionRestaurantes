VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPorDespachar 
   Caption         =   "Cocina en Tiempo Real"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17910
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
   ScaleHeight     =   7380
   ScaleWidth      =   17910
   WindowState     =   2  'Maximized
   Begin VB.Frame FraMesas 
      Caption         =   "Mesa"
      Height          =   1095
      Left            =   6240
      TabIndex        =   23
      Top             =   0
      Width           =   3255
      Begin MSDataListLib.DataCombo DatMesas 
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "Mesas"
      End
   End
   Begin VB.Frame FraZonas 
      Height          =   1050
      Left            =   9650
      TabIndex        =   21
      Top             =   50
      Width           =   2655
      Begin MSComctlLib.ListView lvZonas 
         Height          =   855
         Left            =   30
         TabIndex        =   22
         Top             =   105
         Width           =   2600
         _ExtentX        =   4577
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Rango de Fechas"
      Height          =   1050
      Left            =   17400
      TabIndex        =   16
      Top             =   50
      Visible         =   0   'False
      Width           =   3375
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107347969
         CurrentDate     =   41400
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   1800
         TabIndex        =   18
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   107347969
         CurrentDate     =   41400
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   1800
         TabIndex        =   20
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FraSubFamilias 
      Caption         =   "SubFamilias"
      Height          =   1050
      Left            =   3050
      TabIndex        =   14
      Top             =   50
      Width           =   3135
      Begin MSComctlLib.ListView lvSubFamilia 
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Timer tmrFiltro 
      Enabled         =   0   'False
      Left            =   12720
      Top             =   240
   End
   Begin VB.Frame FraFILTRO 
      Caption         =   "Familias"
      Height          =   1050
      Left            =   120
      TabIndex        =   7
      Top             =   50
      Width           =   2895
      Begin MSComctlLib.ListView lvFamilia 
         Height          =   735
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1296
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame FraEjecucion 
      Caption         =   "Opción de Ejecución"
      Height          =   1050
      Left            =   12360
      TabIndex        =   0
      Top             =   50
      Width           =   4935
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "&Comenzar"
         Height          =   600
         Left            =   3480
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   600
         Left            =   2040
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin MSComCtl2.UpDown udSegundos 
         Height          =   285
         Left            =   3001
         TabIndex        =   4
         Top             =   360
         Width           =   255
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtSegundos"
         BuddyDispid     =   196620
         OrigLeft        =   3600
         OrigTop         =   360
         OrigRight       =   3855
         OrigBottom      =   735
         Max             =   60
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSegundos 
         Height          =   285
         Left            =   2040
         TabIndex        =   3
         Text            =   "5"
         Top             =   360
         Width           =   960
      End
      Begin VB.OptionButton OptAutomatico 
         Caption         =   "Automatico"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OptManual 
         Caption         =   "Manual"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame FraPorDespachar 
      Height          =   6135
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   10335
      Begin MSFlexGridLib.MSFlexGrid MSHPorDespachar 
         Height          =   3495
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6165
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvPorDespachar 
         Height          =   1695
         Left            =   120
         TabIndex        =   12
         Top             =   4320
         Visible         =   0   'False
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblWidth 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ITEMS POR DESPACHAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3675
      End
   End
   Begin VB.Frame FraDespachados 
      Height          =   6135
      Left            =   10560
      TabIndex        =   9
      Top             =   1080
      Width           =   6495
      Begin MSComctlLib.ListView lvTotales 
         Height          =   5535
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   9763
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALES POR DESPACHAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmPorDespachar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cCantidad As Integer
Private cInicio As Boolean

Private Sub Sonar()
If Me.MSHPorDespachar.rows - 1 > cCantidad Then
    cCantidad = Me.MSHPorDespachar.rows - 1
    PlaySound "C:\Admin\Sonidos\alarm.wav.", 1, 1
End If
End Sub

Private Sub RealizarBusqueda(Optional xResumen As Boolean = False)
    cCantidad = Me.MSHPorDespachar.rows - 1
    Me.lvPorDespachar.ListItems.Clear
    Me.lvTotales.ListItems.Clear
    
    Dim xITEM As Object

    Dim sSF, sF  As String

    sSF = ""
    sF = ""
    
    For Each xITEM In Me.lvFamilia.ListItems

        If xITEM.Checked Then
            If CStr(xITEM.Tag) <> "-1" Then
                sF = sF + CStr(xITEM.Tag) & ","
            End If
        End If

    Next

    For Each xITEM In Me.lvSubFamilia.ListItems

        If xITEM.Checked Then
            sSF = sSF + Split(CStr(xITEM.Tag), "-")(0) & ","
        End If

    Next
    
    Dim sZON As String

    sZON = ""

    For Each xITEM In Me.lvZonas.ListItems

        If xITEM.Checked Then
            sZON = sZON + CStr(xITEM.Tag) & ","
        End If

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPITEMSxDESPACHAR"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFAMILIA", adBigInt, adParamInput, , Me.DatFamilia.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ZONA", adVarChar, adParamInput, 800, sZON)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MESA", adVarChar, adParamInput, 10, Me.DatMesas.BoundText)

    If Len(Trim(sF)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FAMILIA", adVarChar, adParamInput, 4000, sF)
    End If

    If Len(Trim(sSF)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SUBFAMILIAS", adVarChar, adParamInput, 4000, sSF)
    End If

    Dim ORSDatos As ADODB.Recordset

    Set ORSDatos = oCmdEjec.Execute

    Dim oRSr  As ADODB.Recordset

    Dim ITEMd As Object
    
    With MSHPorDespachar
        ' deshabilita el repintado para que sea mas rápido
        .Redraw = False
        'Cantidad de filas y columnas
        .rows = 1
        .Cols = ORSDatos.Fields.count
    End With

    For c = 0 To ORSDatos.Fields.count - 1

        'Añade el título del campo al encabezado de columna
        If c = 0 Then
            MSHPorDespachar.TextMatrix(0, c) = ""
        Else
            MSHPorDespachar.TextMatrix(0, c) = ORSDatos.Fields(c).Name
        End If

    Next c
    
    'titulos de columnas
    MSHPorDespachar.TextMatrix(0, 0) = ""
    MSHPorDespachar.TextMatrix(0, 2) = "CANT"
    
    Me.MSHPorDespachar.ColWidth(1) = 5060
    Me.MSHPorDespachar.ColWidth(2) = 1060
    Me.MSHPorDespachar.ColWidth(9) = 0
    Me.MSHPorDespachar.ColWidth(10) = 0
    Me.MSHPorDespachar.ColWidth(11) = 0
    Me.MSHPorDespachar.ColWidth(12) = 0
    Me.MSHPorDespachar.ColWidth(13) = 0
    Me.MSHPorDespachar.ColWidth(14) = 0
    
    Dim fila  As Integer

    Dim xMarc As Boolean

    xMarc = False
    fila = 1

    Do While Not ORSDatos.EOF
        ' Añade una nueva fila
        MSHPorDespachar.rows = MSHPorDespachar.rows + 1

        For c = 0 To ORSDatos.Fields.count - 1
           
            ' si la columna es el campo de tipo CheckBox ...
            If c = 0 Then

                With MSHPorDespachar
                    .Row = fila ' se posiciona en la fila
                    .COL = c '  .. en la columna
                    ' cambia la fuente para esta celda
                    .CellFontName = "Wingdings"
                    .CellFontSize = 14
                    .CellAlignment = flexAlignCenterCenter

                    ' edita la celda
                    If CBool(ORSDatos(0).Value) = True Then
                        .TextMatrix(fila, 0) = Chr(254) ' false
                        .CellForeColor = vbBlue 'The color you want
                        xMarc = True
                    Else
                        .TextMatrix(fila, 0) = Chr(168) ' true
                        xMarc = False
                    End If

                End With
                      
            Else

                'Agrega el registro en la fila y columna específica
                
                MSHPorDespachar.COL = c

                If c = 1 Then

                    Dim xfil As Integer

                    Dim x    As String

                    xfil = 1

                    Dim xarray() As String

                    MSHPorDespachar.TextMatrix(fila, c) = ORSDatos.Fields(c).Value & IIf(Len(Trim(ORSDatos.Fields(14).Value)) <> 0, vbCrLf & ORSDatos.Fields(14).Value, "")
                    Me.lblWidth.Caption = ORSDatos.Fields(c).Value & IIf(Len(Trim(ORSDatos.Fields(14).Value)) <> 0, vbCrLf & ORSDatos.Fields(14).Value, "")

                    If Len(Trim(ORSDatos.Fields(14).Value)) <> 0 Then
                        x = Len(Trim(ORSDatos.Fields(14).Value)) / 30
                        xarray = Split(x, ".")

                        If UBound(xarray) <> 0 Then
                            If xarray(1) <> 0 Then xfil = xfil + xarray(0) + 1
                        End If
                    End If

                    MSHPorDespachar.RowHeight(fila) = xfil * 315
                Else
                    MSHPorDespachar.TextMatrix(fila, c) = ORSDatos.Fields(c).Value
                End If
            End If
            
            'MSHPorDespachar.ColWidth(0) = lblWidth.Width + 100
            'InMSHPorDespachar.
           
            If xMarc Then MSHPorDespachar.CellForeColor = vbBlue
            
        Next
            
        ' Siguiente registro
        ORSDatos.MoveNext
        fila = fila + 1 'Incrementa la posición de la fila actual
    Loop

    ' If Me.MSHPorDespachar.Rows > 1 Then
    ' Me.MSHPorDespachar.Row = 2
    '            End If
            
    Me.MSHPorDespachar.ColWidth(0) = 400
    MSHPorDespachar.Redraw = True

    'If Me.MSHPorDespachar.Rows <> 1 Then
    '    cCantidad = Me.MSHPorDespachar.Rows
    'End If
    If ORSDatos.RecordCount <> 0 Then ORSDatos.MoveFirst
   
    If ORSDatos.RecordCount <> 0 Then
    
        Sonar
        'cCantidad = ORSDatos.RecordCount
    End If

    Set oRSr = ORSDatos.NextRecordset

    Dim i As Integer

    For i = 1 To Me.lvPorDespachar.ListItems.count
        Me.lvPorDespachar.ListItems(i).Selected = False
    
    Next

    Do While Not oRSr.EOF
        Set ITEMd = Me.lvTotales.ListItems.Add(, , oRSr!producto)
        ITEMd.SubItems(1) = oRSr!Cantidad
        ITEMd.Tag = oRSr!Codigo
        oRSr.MoveNext
    Loop

End Sub

Private Sub cmdActualizar_Click()

    Dim ITEMf As Object

    Dim xMarc As Boolean

    xMarc = False

    For Each ITEMf In Me.lvFamilia.ListItems

        If ITEMf.Checked Then
            xMarc = True

            Exit For

        End If

    Next
    
    If Not xMarc Then
        MsgBox "Debe marcar una Familia"
    Else
        RealizarBusqueda
    End If
    
End Sub

Private Sub cmdEjecutar_Click()

    If Me.cmdEjecutar.Caption = "&Comenzar" Then
        Me.tmrFiltro.Interval = Me.txtSegundos.Text * 1000
        Me.tmrFiltro.Enabled = True
        Me.cmdEjecutar.Caption = "&Detener"
    Else
        Me.cmdEjecutar.Caption = "&Comenzar"
        Me.tmrFiltro.Enabled = False
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF6 Then
' MsgBox Me.MSHPorDespachar.TextMatrix(Me.MSHPorDespachar.MouseRow, 1)
'End If
End Sub

Private Sub Form_Load()
    ConfigurarLV
    cInicio = True
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpListarFamilias2"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    
    Dim ORSf As ADODB.Recordset

    Set ORSf = oCmdEjec.Execute
    
'    Set Me.DatFamilia.RowSource = ORSf
'    Me.DatFamilia.BoundColumn = ORSf.Fields(0).Name
'    Me.DatFamilia.ListField = ORSf.Fields(1).Name
'    Me.DatFamilia.BoundText = -1
'
    
    Dim itemX As Object

    Do While Not ORSf.EOF
        Set itemX = Me.lvFamilia.ListItems.Add(, , ORSf!Familia)
        itemX.Tag = ORSf!NUMERO
        ORSf.MoveNext
    Loop
    
    For Each itemX In Me.lvFamilia.ListItems
    If CStr(itemX.Tag) = "-1" Then
        itemX.Checked = True
        Exit For
    End If
        
    Next
    
    
    
    Me.dtpDesde.Value = LK_FECHA_DIA
    Me.dtpHasta.Value = LK_FECHA_DIA

    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpListarZonas"
    Set ORSf = oCmdEjec.Execute(, LK_CODCIA)
    
    Me.lvZonas.ListItems.Clear
    
    

    Set itemX = Me.lvZonas.ListItems.Add(, , ".: TODOS :.")
    itemX.Tag = -1

    Do While Not ORSf.EOF
        Set itemX = Me.lvZonas.ListItems.Add(, , ORSf!denomina)
        itemX.Tag = ORSf!Codigo
        ORSf.MoveNext
    Loop
    
    Me.lvZonas.ListItems(1).Checked = True
    
    'COMBONUEVO
    Dim orsM As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPORDESPACHAR_LISTARMESAS"
    Set orsM = oCmdEjec.Execute(, LK_CODCIA)
    Set Me.DatMesas.RowSource = orsM
    Me.DatMesas.BoundColumn = orsM.Fields(0).Name
    Me.DatMesas.ListField = orsM.Fields(1).Name
    Me.DatMesas.BoundText = -1
End Sub

Private Sub ConfigurarLV()

    With MSHPorDespachar
        .rows = 2
        .Cols = 6
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionFree
    End With
    
    With Me.lvPorDespachar
        .ColumnHeaders.Add , , "PRODUCTO", 5000
        .ColumnHeaders.Add , , "CANT.", 1200
        .ColumnHeaders.Add , , "MESA", 1200
        .ColumnHeaders.Add , , "MOZO"
        .ColumnHeaders.Add , , "HORA", 1640
        .ColumnHeaders.Add , , "TIEMPO", 1640
        .ColumnHeaders.Add , , "TIEMPO PREPARACION", 1640
        .ColumnHeaders.Add , , "DETALLE"
        .ColumnHeaders.Add , , "NRO", 0
        .ColumnHeaders.Add , , "SEC", 0
        .ColumnHeaders.Add , , "SERIE", 0
        .ColumnHeaders.Add , , "CODART", 0
        '.ColumnHeaders.Add , , "FECHA", 0
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
    End With

    With Me.lvTotales
        .ColumnHeaders.Add , , "PRODUCTO", 3800
        .ColumnHeaders.Add , , "CANT.", 1200
        .FullRowSelect = True
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
    End With
    
    With Me.lvSubFamilia
        .ColumnHeaders.Add , , "subfamilia", Me.lvSubFamilia.Width - 400
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .View = lvwReport
        .CheckBoxes = True
    End With
    
      With Me.lvFamilia
        .ColumnHeaders.Add , , "Familia", Me.lvFamilia.Width - 400
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .View = lvwReport
        .CheckBoxes = True
    End With
    
    With Me.lvZonas
        .ColumnHeaders.Add , , "ZONA", Me.lvZonas.Width - 400
        .FullRowSelect = True
        .Gridlines = False
        .LabelEdit = lvwManual
        .View = lvwReport
        .CheckBoxes = True
    End With

End Sub

Private Sub Form_Resize()

    If (Me.ScaleWidth - 6300) <= 0 Then Exit Sub
    If (Me.ScaleHeight - 1800) <= 0 Then Exit Sub
    Me.FraPorDespachar.Width = Me.ScaleWidth - 6000
    Me.FraDespachados.Left = Me.FraPorDespachar.Width + 200
    Me.FraDespachados.Width = Me.ScaleWidth - (350 + Me.FraPorDespachar.Width)
    'listview
    Me.lvPorDespachar.Width = Me.ScaleWidth - 6300
    Me.lvTotales.Width = Me.ScaleWidth - (630 + Me.FraPorDespachar.Width)
    'ALTO
    Me.FraPorDespachar.Height = Me.ScaleHeight - 1200
    Me.FraDespachados.Height = Me.FraPorDespachar.Height
    Me.lvPorDespachar.Height = Me.ScaleHeight - 1800
    Me.lvTotales.Height = Me.lvPorDespachar.Height
    
     Me.MSHPorDespachar.Width = Me.ScaleWidth - 6300
    Me.MSHPorDespachar.Height = Me.ScaleHeight - 1800
End Sub

Private Sub lvFamilia_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim itemX As Object

    If Item.Checked Then 'marcado
        If CStr(Item.Tag) = "-1" Then
            Me.lvSubFamilia.ListItems.Clear

            For Each itemX In Me.lvFamilia.ListItems

                If CStr(itemX.Tag) <> "-1" Then
                    itemX.Checked = False
                End If

            Next

        Else
   
            For Each itemX In Me.lvFamilia.ListItems

                If CStr(itemX.Tag) = "-1" Then
                    itemX.Checked = False
                    Exit For
                End If

            Next

            LimpiaParametros oCmdEjec
            oCmdEjec.CommandText = "SpListarSubFamiliasxFamilia"
            oCmdEjec.CommandType = adCmdStoredProc

            Dim ORSsf As ADODB.Recordset

            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@idfamilia", adBigInt, adParamInput, , Item.Tag)

            Set ORSsf = oCmdEjec.Execute
            'Me.lvSubFamilia.ListItems.Clear

            Do While Not ORSsf.EOF
                Set itemX = Me.lvSubFamilia.ListItems.Add(, , ORSsf!subfamilia)
                itemX.Tag = ORSsf!NUMERO & "-" & Item.Tag
                ORSsf.MoveNext
            Loop

        End If

    Else 'desmarcado

        For i = Me.lvSubFamilia.ListItems.count To 1 Step -1

            If Split(CStr(Me.lvSubFamilia.ListItems(i).Tag), "-")(1) = CStr(Item.Tag) Then
                Me.lvSubFamilia.ListItems.Remove i
            End If
           
        Next

    End If

End Sub

Private Sub lvPorDespachar_DblClick()
If Me.txtSegundos.Visible Then  'ESTA EJECUTANDO
    Me.tmrFiltro.Enabled = False
    
End If
    frmPorDespacharCantidad.vCANTIDAD = Me.lvPorDespachar.SelectedItem.SubItems(1)
    frmPorDespacharCantidad.lblproducto.Caption = Me.lvPorDespachar.SelectedItem.Text
    frmPorDespacharCantidad.Show vbModal

    If frmPorDespacharCantidad.vDespachado Then

       ' If MsgBox("¿Confirma Despacho?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub   quitado gts
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCOMANDADESPACHAR"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.lvPorDespachar.SelectedItem.SubItems(10))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.lvPorDespachar.SelectedItem.SubItems(8))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , Me.lvPorDespachar.SelectedItem.SubItems(9))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODART", adBigInt, adParamInput, , Me.lvPorDespachar.SelectedItem.SubItems(11))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, , frmPorDespacharCantidad.vCANTIDAD)
        oCmdEjec.Execute
    
        cmdActualizar_Click
    End If
If Me.txtSegundos.Visible = True Then
    Me.tmrFiltro.Enabled = True
    RealizarBusqueda
End If
    
    
End Sub

Private Sub lvPorDespachar_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo Exito

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPCANTAPEDIDO"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adDouble, adParamInput, , Item.SubItems(11))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adDouble, adParamInput, , Item.SubItems(9))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adDouble, adParamInput, , Item.SubItems(8))
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTADO", adBoolean, adParamInput, 1, IIf(Item.Checked, True, False))
    
    
    oCmdEjec.Execute
    
    If Not Item.Checked Then
        Me.lvPorDespachar.ListItems(Item.index).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(1).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(2).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(3).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(4).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(5).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(6).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(7).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(8).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(9).ForeColor = vbBlack
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(10).ForeColor = vbBlack
   

    Else
        Me.lvPorDespachar.ListItems(Item.index).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(1).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(2).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(3).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(4).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(5).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(6).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(7).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(8).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(9).ForeColor = vbBlue
        Me.lvPorDespachar.ListItems(Item.index).ListSubItems(10).ForeColor = vbBlue

    End If

   ' MsgBox "Datos Modificados correctamente."
        
    Exit Sub

Exito:
    MsgBox Err.Description
End Sub

Private Sub lvTotales_DblClick()
'MsgBox lvTotales.SelectedItem.Tag
RealizarBusquedaFiltrada Me.lvTotales.SelectedItem.Tag
End Sub

Private Sub lvZonas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Tag = -1 Then
If Item.Checked = False Then Item.Checked = True
    Dim i As Integer
    For i = 1 To Me.lvZonas.ListItems.count
        If Me.lvZonas.ListItems(i).Tag <> -1 Then
        Me.lvZonas.ListItems(i).Checked = False
        End If
    Next
    Else
    Me.lvZonas.ListItems(1).Checked = False
End If
End Sub

Private Sub MSHPorDespachar_Click()

    Dim c     As Integer

    Dim xFILA As Integer

    Dim xCOL  As Integer

    xFILA = MSHPorDespachar.MouseRow
    xCOL = MSHPorDespachar.MouseCol

    With MSHPorDespachar

       
        If (xCOL) = 0 And xFILA <> 0 Then
          
           LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCANTAPEDIDO"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , Me.MSHPorDespachar.TextMatrix(xFILA, 12))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , Me.MSHPorDespachar.TextMatrix(xFILA, 10))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adDouble, adParamInput, , Me.MSHPorDespachar.TextMatrix(xFILA, 9))
    
    
            ' CheckBox en false
            If .TextMatrix(xFILA, 0) = Chr(168) Then 'chekc marcado
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTADO", adBoolean, adParamInput, 1, True)
                .TextMatrix(xFILA, 0) = Chr(254)
               
                For c = 0 To MSHPorDespachar.Cols - 1
                    .COL = c 'The Position (Col)

                    .CellForeColor = vbBlue
                Next
                oCmdEjec.Execute
                Dim cNot As Boolean
                cNot = Leer_Ini(App.Path & "\config.ini", "NOTIFICACION", "0")
                If cNot Then
                EnviarNotificacion Me.MSHPorDespachar.TextMatrix(xFILA, 12), Me.MSHPorDespachar.TextMatrix(xFILA, 10), Me.MSHPorDespachar.TextMatrix(xFILA, 9)
                End If
                ' CheckBox en true
            Else 'check sin marcar
                oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTADO", adBoolean, adParamInput, 1, False)
                .TextMatrix(xFILA, 0) = Chr(168)
                .COL = 0 'The Position (Col)

                For c = 0 To MSHPorDespachar.Cols - 1
                    .COL = c
                    .CellForeColor = vbBlack
                Next
            oCmdEjec.Execute
            End If
            
              cmdActualizar_Click
        End If
    
    End With

End Sub

Private Sub EnviarNotificacion(cIdproducto As Double, cNumSec As Integer, cNumFac As Double)
    Dim p          As Object

    Dim Texto      As String

    Dim sInputJson As String

    Dim cab        As Integer

    Set httpURL = New WinHttp.WinHttpRequest
    
    'obteniendo datos
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "[dbo].[USP_DATOS_NOTIFICACION]"
            oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , cNumSec)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adDouble, adParamInput, , cNumFac)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adDouble, adParamInput, , cIdproducto)
        
        Set oRSmain = oCmdEjec.Execute
    

    Dim cTitulo, cMensaje As String
    If Not oRSmain.EOF Then
        cTitulo = oRSmain!mesa
        cMensaje = oRSmain!cant & Space(2) & oRSmain!Prod
    End If
    

    cadena = "https://jmendozasoft.000webhostapp.com/stodomingo/notify.php?TITULO=" & cTitulo & "&MENSAJE=" & cMensaje
    'cadena = "https://gtsoftwaresac.com/Stodomingo/notify.php?TITULO=" & cTitulo & "&MENSAJE=" & cMensaje
    
    httpURL.Open "GET", cadena
    httpURL.Send
    Texto = httpURL.ResponseText

    If Texto = "[]" Then
        MsgBox ("No se obtuvo resultados")

        Exit Sub

    End If

    sInputJson = "{items:" & "[" & Texto & "]" & "}"

    Set p = JSON.parse(sInputJson)

On Error GoTo Notificacion
    If p.Item("items").Item(1).Item("CODIGO") <> 0 Then
'        MsgBox ("Notificación enviada correctamente")
'    Else
        MsgBox ("Error al enviar notificacion")
    End If
    Exit Sub
Notificacion:
    
End Sub

Private Sub MSHPorDespachar_DblClick()
Dim xFILA As Integer
Dim xCOL As Integer

xFILA = Me.MSHPorDespachar.MouseRow


    If Me.txtSegundos.Visible Then  'ESTA EJECUTANDO
        Me.tmrFiltro.Enabled = False
    
    End If

    frmPorDespacharCantidad.vCANTIDAD = Me.MSHPorDespachar.TextMatrix(xFILA, 2) ' Me.lvPorDespachar.SelectedItem.SubItems(1)
    frmPorDespacharCantidad.lblproducto.Caption = Me.MSHPorDespachar.TextMatrix(xFILA, 1) ' Me.lvPorDespachar.SelectedItem.Text
    frmPorDespacharCantidad.Show vbModal

    If frmPorDespacharCantidad.vDespachado Then

        ' If MsgBox("¿Confirma Despacho?.", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub   quitado gts
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCOMANDADESPACHAR"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
        'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSER", adChar, adParamInput, 3, Me.MSHPorDespachar.TextMatrix(xFILA, 11))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMFAC", adBigInt, adParamInput, , Me.MSHPorDespachar.TextMatrix(xFILA, 9))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@NUMSEC", adInteger, adParamInput, , Me.MSHPorDespachar.TextMatrix(xFILA, 10))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODART", adBigInt, adParamInput, , Me.MSHPorDespachar.TextMatrix(xFILA, 12))
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CANTIDAD", adDouble, adParamInput, , frmPorDespacharCantidad.vCANTIDAD)
        oCmdEjec.Execute
    cCantidad = cCantidad - 1
        cmdActualizar_Click
    End If

    If Me.txtSegundos.Visible = True Then
        Me.tmrFiltro.Enabled = True
        
        RealizarBusqueda
        
    End If

End Sub

Private Sub OptAutomatico_Click()
Me.cmdActualizar.Visible = False
Me.cmdEjecutar.Visible = True
Me.tmrFiltro.Enabled = True
Me.txtSegundos.Visible = True
Me.cmdEjecutar.Caption = "&Comenzar"
End Sub

Private Sub OptManual_Click()
Me.cmdActualizar.Visible = True
Me.cmdEjecutar.Visible = False
Me.txtSegundos.Visible = False
Me.tmrFiltro.Enabled = False
End Sub

Private Sub tmrFiltro_Timer()
RealizarBusqueda
End Sub

Private Sub RealizarBusquedaFiltrada(xIDproducto As Double)
    Me.lvPorDespachar.ListItems.Clear
    'Me.lvTotales.ListItems.Clear
    
    Dim xITEM As Object

    Dim sSF, sF  As String

    sSF = ""
    sF = ""

    For Each xITEM In Me.lvFamilia.ListItems

        If xITEM.Checked Then
            If CStr(xITEM.Tag) <> "-1" Then
                sF = sF + CStr(xITEM.Tag) & ","
            End If
        End If

    Next
    
    For Each xITEM In Me.lvSubFamilia.ListItems

        If xITEM.Checked Then
            sSF = sSF + CStr(xITEM.Tag) & ","
        End If

    Next
    
    Dim sZON As String

    sZON = ""

    For Each xITEM In Me.lvZonas.ListItems

        If xITEM.Checked Then
            sZON = sZON + CStr(xITEM.Tag) & ","
        End If

    Next

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPITEMSxDESPACHAR2"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CodCia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA1", adDBTimeStamp, adParamInput, , Me.dtpDesde.Value)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA2", adDBTimeStamp, adParamInput, , Me.dtpHasta.Value)
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDFAMILIA", adBigInt, adParamInput, , Me.DatFamilia.BoundText)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ZONA", adVarChar, adParamInput, 800, sZON)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPRODUCTO", adBigInt, adParamInput, , xIDproducto)

    If Len(Trim(sF)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FAMILIA", adVarChar, adParamInput, 4000, sF)
    End If

    If Len(Trim(sSF)) <> 0 Then
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@SUBFAMILIAS", adVarChar, adParamInput, 4000, sSF)
    End If

    Dim ORSDatos As ADODB.Recordset

    Set ORSDatos = oCmdEjec.Execute

    Dim oRSr As ADODB.Recordset
    
    'Dim ORSDatos As ADODB.Recordset

    Set ORSDatos = oCmdEjec.Execute

    'Dim oRSr  As ADODB.Recordset

    Dim ITEMd As Object
    
    With MSHPorDespachar
        ' deshabilita el repintado para que sea mas rápido
        .Redraw = False
        'Cantidad de filas y columnas
        .rows = 1
        .Cols = ORSDatos.Fields.count
    End With

    For c = 0 To ORSDatos.Fields.count - 1

        'Añade el título del campo al encabezado de columna
        If c = 0 Then
            MSHPorDespachar.TextMatrix(0, c) = ""
        Else
            MSHPorDespachar.TextMatrix(0, c) = ORSDatos.Fields(c).Name
        End If

    Next c
    
    'titulos de columnas
    MSHPorDespachar.TextMatrix(0, 0) = ""
    MSHPorDespachar.TextMatrix(0, 2) = "CANT"
    
    Me.MSHPorDespachar.ColWidth(1) = 5060
    Me.MSHPorDespachar.ColWidth(2) = 1060
    Me.MSHPorDespachar.ColWidth(9) = 0
    Me.MSHPorDespachar.ColWidth(10) = 0
    Me.MSHPorDespachar.ColWidth(11) = 0
    Me.MSHPorDespachar.ColWidth(12) = 0
    Me.MSHPorDespachar.ColWidth(13) = 0
    Me.MSHPorDespachar.ColWidth(14) = 0
    
    Dim fila  As Integer

    Dim xMarc As Boolean

    xMarc = False
    fila = 1

    Do While Not ORSDatos.EOF
        ' Añade una nueva fila
        MSHPorDespachar.rows = MSHPorDespachar.rows + 1

        For c = 0 To ORSDatos.Fields.count - 1
           
            ' si la columna es el campo de tipo CheckBox ...
            If c = 0 Then

                With MSHPorDespachar
                    .Row = fila ' se posiciona en la fila
                    .COL = c '  .. en la columna
                    ' cambia la fuente para esta celda
                    .CellFontName = "Wingdings"
                    .CellFontSize = 14
                    .CellAlignment = flexAlignCenterCenter

                    ' edita la celda
                    If CBool(ORSDatos(0).Value) = True Then
                        .TextMatrix(fila, 0) = Chr(254) ' false
                        .CellForeColor = vbBlue 'The color you want
                        xMarc = True
                    Else
                        .TextMatrix(fila, 0) = Chr(168) ' true
                        xMarc = False
                    End If

                End With
                      
            Else

                'Agrega el registro en la fila y columna específica
                
                MSHPorDespachar.COL = c

                If c = 1 Then

                    Dim xfil As Integer

                    Dim x    As String

                    xfil = 1

                    Dim xarray() As String

                    MSHPorDespachar.TextMatrix(fila, c) = ORSDatos.Fields(c).Value & IIf(Len(Trim(ORSDatos.Fields(14).Value)) <> 0, vbCrLf & ORSDatos.Fields(14).Value, "")
                    Me.lblWidth.Caption = ORSDatos.Fields(c).Value & IIf(Len(Trim(ORSDatos.Fields(14).Value)) <> 0, vbCrLf & ORSDatos.Fields(14).Value, "")

                    If Len(Trim(ORSDatos.Fields(14).Value)) <> 0 Then
                        x = Len(Trim(ORSDatos.Fields(14).Value)) / 30
                        xarray = Split(x, ".")

                        If UBound(xarray) <> 0 Then
                            If xarray(1) <> 0 Then xfil = xfil + xarray(0) + 1
                        End If
                    End If

                    MSHPorDespachar.RowHeight(fila) = xfil * 315
                Else
                    MSHPorDespachar.TextMatrix(fila, c) = ORSDatos.Fields(c).Value
                End If
            End If

            If xMarc Then MSHPorDespachar.CellForeColor = vbBlue
            
        Next

        ' Siguiente registro
        ORSDatos.MoveNext
        fila = fila + 1 'Incrementa la posición de la fila actual
    Loop

    Me.MSHPorDespachar.ColWidth(0) = 400
    MSHPorDespachar.Redraw = True

    If ORSDatos.RecordCount <> 0 Then ORSDatos.MoveFirst
    If ORSDatos.RecordCount <> 0 Then
    
        Sonar
    End If

    Set oRSr = ORSDatos.NextRecordset

End Sub
