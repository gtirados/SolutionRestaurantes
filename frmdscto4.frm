VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F6E4F630-E903-11D5-8BB9-0080AD40A177}#1.8#0"; "OSControlsUser.ocx"
Begin VB.Form Frmdscto4 
   Caption         =   "Mantenimiento de Porcentajes de Descuento"
   ClientHeight    =   6210
   ClientLeft      =   330
   ClientTop       =   1380
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin OSControlsUser.OSMessage OSMessage1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   39
      Top             =   5880
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   582
   End
   Begin VB.CommandButton CmdValor 
      Caption         =   "&Ingresar Valor"
      Height          =   450
      Left            =   5490
      TabIndex        =   36
      Top             =   3945
      Width           =   1935
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9180
      TabIndex        =   27
      Top             =   3885
      Width           =   1305
   End
   Begin VB.CommandButton Cmdsalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10530
      TabIndex        =   26
      Top             =   3870
      Width           =   1305
   End
   Begin VB.ComboBox Art_Clase 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8205
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   900
      Width           =   3390
   End
   Begin VB.ComboBox Art_Marca 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4305
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   900
      Width           =   3390
   End
   Begin VB.ComboBox art_familia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   405
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   210
      Width           =   3390
   End
   Begin VB.ComboBox art_subfam 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4305
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   210
      Width           =   3390
   End
   Begin VB.ComboBox art_grupo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8205
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   225
      Width           =   3390
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   10080
      TabIndex        =   8
      Top             =   1980
      Width           =   1860
      Begin VB.CommandButton Command1 
         Caption         =   "Fi&ltro Seleccion"
         Height          =   390
         Left            =   120
         TabIndex        =   38
         Top             =   165
         Width           =   1650
      End
      Begin VB.CommandButton Cmdfiltro 
         Caption         =   "Filtro &Total"
         Height          =   390
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1650
      End
   End
   Begin VB.TextBox TxtArt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   915
      TabIndex        =   4
      Top             =   3870
      Width           =   4200
   End
   Begin VB.Frame vista 
      Caption         =   "Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1905
      Left            =   45
      TabIndex        =   3
      Top             =   1905
      Width           =   5175
      Begin MSComctlLib.ListView VART 
         Height          =   1620
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2858
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   128
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ITEM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPCION"
            Object.Width           =   7832
         EndProperty
      End
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   75
      TabIndex        =   0
      Top             =   1545
      Width           =   3450
   End
   Begin OSControlsUser.OSNewFlexGrid Prec 
      Height          =   2280
      Left            =   135
      TabIndex        =   2
      Top             =   4470
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   4022
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      BorderStyle     =   0
      BackColorSel    =   8388608
      Cols            =   12
      ColEdit(4)      =   0   'False
      ColSalto(4)     =   0
      ColAtras(4)     =   0
      FixedCols       =   4
      ColWidth0       =   1020
   End
   Begin VB.Frame Frame3 
      Height          =   1905
      Left            =   5220
      TabIndex        =   5
      Top             =   1905
      Width           =   4785
      Begin VB.ListBox LstDes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   1605
         Left            =   2445
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   210
         Width           =   2205
      End
      Begin VB.ListBox LstLis 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   1605
         Left            =   150
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   210
         Width           =   2220
      End
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   240
      Left            =   4545
      TabIndex        =   24
      Top             =   1380
      Visible         =   0   'False
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Max             =   10000
      Scrolling       =   1
   End
   Begin VB.ComboBox art_numero 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   405
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   900
      Width           =   3390
   End
   Begin VB.Frame Mante 
      Caption         =   "Mantenimiento de descuentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2580
      Left            =   5235
      TabIndex        =   25
      Top             =   1860
      Visible         =   0   'False
      Width           =   6660
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1815
         TabIndex        =   35
         Top             =   1620
         Width           =   3210
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1815
         TabIndex        =   34
         Top             =   1035
         Width           =   2505
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1830
         TabIndex        =   33
         Top             =   540
         Width           =   1320
      End
      Begin VB.CommandButton CmdGraba 
         Caption         =   "&Nuevo"
         Height          =   495
         Left            =   5160
         TabIndex        =   29
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "C&ancelar"
         Height          =   495
         Left            =   5160
         TabIndex        =   28
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Largo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   10
         Left            =   210
         TabIndex        =   32
         Top             =   1620
         Width           =   1320
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Corto :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   9
         Left            =   210
         TabIndex        =   31
         Top             =   1110
         Width           =   1290
      End
      Begin VB.Label lblart 
         AutoSize        =   -1  'True
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   30
         Top             =   630
         Width           =   600
      End
   End
   Begin VB.Label lblporcentaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0%..."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6450
      TabIndex        =   37
      Top             =   1605
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "Clase:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   7
      Left            =   8205
      TabIndex        =   23
      Top             =   690
      Width           =   540
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   6
      Left            =   4275
      TabIndex        =   21
      Top             =   690
      Width           =   600
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "Articulo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   3870
      Width           =   720
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "Sigla. Art:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   75
      TabIndex        =   14
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "Sub Linea:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   3
      Left            =   375
      TabIndex        =   13
      Top             =   690
      Width           =   930
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "División:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   0
      Left            =   375
      TabIndex        =   12
      Top             =   0
      Width           =   750
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "Familia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   1
      Left            =   4275
      TabIndex        =   11
      Top             =   0
      Width           =   660
   End
   Begin VB.Label lblart 
      AutoSize        =   -1  'True
      Caption         =   "Linea:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   2
      Left            =   8205
      TabIndex        =   10
      Top             =   0
      Width           =   540
   End
End
Attribute VB_Name = "Frmdscto4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cnp As New ADODB.Connection
Public TL As New ADODB.Recordset
Public Pr As New ADODB.Recordset
Public Ar As New ADODB.Recordset
Public DS As New ADODB.Recordset
Public Td As New ADODB.Recordset
Public ART As New ADODB.Recordset
Dim mItem As MSComctlLib.ListItem
' Dim MItem As ListItem
'''''''''''''''''''''''''''''
Dim RsC As New ADODB.Recordset
Dim CmdFamilias As New ADODB.Command
Dim PrmCodCia As New ADODB.Parameter
Dim PrmTipReg As New ADODB.Parameter
Dim PrmCodArt As New ADODB.Parameter
''''''''''''''''''''''''
''Filtrando tabla articulos
Dim RsArt As New ADODB.Recordset
Dim CmdIp As New ADODB.Command
Dim PrmPack As New ADODB.Parameter
Dim PrmDiv As New ADODB.Parameter
Dim PrmFam As New ADODB.Parameter
Dim PrmLinea As New ADODB.Parameter
Dim PrmSublinea As New ADODB.Parameter
Dim PrmMarca As New ADODB.Parameter
Dim PrmClase As New ADODB.Parameter
Dim PrmDesc As New ADODB.Parameter
Dim PrmAlterno As New ADODB.Parameter
''''''''''''''''''''''''
Dim I As Long
Dim J As Long
Dim F As Long
Dim C As Long
''''''''''''''''''''
Dim VALOR As Single
Dim X() As String
Dim Xx() As String
Public Compa As String
Dim Codigo As String
Dim VDes As Integer
Dim VLis As Integer
Dim SQL As String

'''''''''''''''''''''''
Sub set_Cmd()
  With CmdFamilias
      .ActiveConnection = cnp
      .CommandType = adCmdStoredProc
      .CommandText = "Sp_List_Familias"
      '''''definiendo parametros para el comando de busca familias
        PrmCodCia.Type = adChar: PrmCodCia.Size = 2: PrmCodCia.Direction = adParamInput
        PrmTipReg.Type = adInteger: PrmTipReg.Direction = adParamInput
            ''''o
      With PrmCodArt
            .Type = adInteger: .Direction = adParamInput
      End With
      .Parameters.Append PrmCodCia
      .Parameters.Append PrmTipReg
      .Parameters.Append PrmCodArt
  End With
  '''commando de filtro articulos
  With CmdIp
         .ActiveConnection = cnp
         .CommandType = adCmdStoredProc
         .CommandText = "Sp_List_Pack"
          
          PrmPack.Type = adChar: PrmPack.Size = 1: PrmPack.Direction = adParamInput
          PrmDiv.Type = adChar: PrmDiv.Size = 4: PrmDiv.Direction = adParamInput
          PrmFam.Type = adChar: PrmFam.Size = 4: PrmFam.Direction = adParamInput
          PrmLinea.Type = adChar: PrmLinea.Size = 4: PrmLinea.Direction = adParamInput
          PrmSublinea.Type = adChar: PrmSublinea.Size = 4: PrmSublinea.Direction = adParamInput
          PrmMarca.Type = adChar: PrmMarca.Size = 4: PrmMarca.Direction = adParamInput
          PrmClase.Type = adChar: PrmClase.Size = 4: PrmClase.Direction = adParamInput
          PrmDesc.Type = adChar: PrmDesc.Size = 40: PrmDesc.Direction = adParamInput
          PrmAlterno.Type = adChar: PrmAlterno.Size = 10: PrmAlterno.Direction = adParamInput
          .Parameters.Append PrmCodCia
          .Parameters.Append PrmPack
          .Parameters.Append PrmDiv
          .Parameters.Append PrmFam
          .Parameters.Append PrmLinea
          .Parameters.Append PrmSublinea
          .Parameters.Append PrmMarca
          .Parameters.Append PrmClase
          .Parameters.Append PrmDesc
          .Parameters.Append PrmAlterno
  End With
  
End Sub
Sub Load_Familias(CboControl As ComboBox)
''''Procedimiento para llenar Familias,sub familias,etc
On Error GoTo ErrNuevo
'Ejecutando el comando
Set RsC = CmdFamilias.Execute
CboControl.Clear
If RsC.RecordCount > 0 Then
  RsC.MoveFirst
End If

While Not RsC.EOF
    With CboControl
       .AddItem StrConv(Trim(RsC!tab_NOMLARGO), vbProperCase)
       .ItemData(.NewIndex) = RsC!TAB_NUMTAB
       RsC.MoveNext
    End With
    
Wend
Exit Sub
ErrNuevo:
 Set RsC = Nothing
 MsgBox Err.Description
 
End Sub
Sub llenaArt()
Dim Clave As String
Set RsArt = CmdIp.Execute
VART.ListItems.Clear
If RsArt.RecordCount > 0 Then
 RsArt.MoveFirst
End If

With RsArt
    While Not .EOF
             Clave = !art_KEY & "A"
             Set mItem = VART.ListItems.Add(Key:=Clave, Text:=!art_KEY)
             mItem.ListSubItems.Add Text:=!art_nombre
             mItem.ListSubItems.Add Text:=!art_KEY
            .MoveNext
          Wend
      End With
End Sub

Private Sub Art_Clase_Click()
'''''''''''''''''''''
If Art_Clase.ListIndex <> -1 Then
 PrmClase = Art_Clase.ItemData(Art_Clase.ListIndex)
End If
 llenaArt
End Sub

Private Sub art_familia_Click()
''''''''''''''limpiando combos
If art_familia.ListIndex <> -1 Then

  art_subfam.ListIndex = -1
  art_grupo.ListIndex = -1
  art_numero.ListIndex = -1
  Art_Marca.ListIndex = -1
  Art_Clase.ListIndex = -1

'''''''''''''''''''
    PrmCodArt = art_familia.ItemData(art_familia.ListIndex)
    PrmTipReg = "123"
    Load_Familias art_subfam
    ''''''''''''''''''
    PrmCodCia = LK_CODCIA
    PrmPack = "M"
    PrmDiv = art_familia.ItemData(art_familia.ListIndex)
    PrmFam = Null
    PrmLinea = Null
    PrmSublinea = Null
    PrmMarca = Null
    PrmClase = Null
    llenaArt
End If
End Sub

Private Sub art_grupo_Click()
''''''''''''''limpiando combos
art_numero.ListIndex = -1
Art_Marca.ListIndex = -1
Art_Clase.ListIndex = -1

'''''''''''''''''''
If art_grupo.ListIndex <> -1 Then
 PrmCodArt = art_grupo.ItemData(art_grupo.ListIndex)
 PrmTipReg = "130"
 Load_Familias art_numero
 PrmLinea = art_grupo.ItemData(art_grupo.ListIndex)
End If
''''''''''''''''''
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
llenaArt

End Sub

Private Sub Art_Marca_Click()
''''''''''''''limpiando combos
Art_Clase.ListIndex = -1
'''''''''''''''''''
If Art_Marca.ListIndex <> -1 Then
    PrmCodArt = Art_Marca.ItemData(Art_Marca.ListIndex)
    PrmTipReg = "132"
    Load_Familias Art_Clase
    PrmMarca = Art_Marca.ItemData(Art_Marca.ListIndex)
 '''''''''''''''''''''
 End If
 PrmClase = Null
 llenaArt
End Sub

Private Sub art_numero_Click()
''''''''''''''limpiando combos
Art_Marca.ListIndex = -1
Art_Clase.ListIndex = -1
'''''''''''''''''''
If art_numero.ListIndex <> -1 Then

    PrmCodArt = art_numero.ItemData(art_numero.ListIndex)
    PrmTipReg = "131"
    Load_Familias Art_Marca
    PrmSublinea = art_numero.ItemData(art_numero.ListIndex)
 End If
 '''''''''''''''''''''
 PrmMarca = Null
 PrmClase = Null
 llenaArt
End Sub
Private Sub art_subfam_Click()
''''''''''''''limpiando combos
  art_grupo.ListIndex = -1
  art_numero.ListIndex = -1
  Art_Marca.ListIndex = -1
  Art_Clase.ListIndex = -1
'''''''''''''''''''
If art_subfam.ListIndex <> -1 Then
  PrmCodArt = art_subfam.ItemData(art_subfam.ListIndex)
  PrmTipReg = "129"
  Load_Familias art_grupo
  PrmFam = art_subfam.ItemData(art_subfam.ListIndex)
End If
''''''''''''''''''
PrmLinea = Null
PrmSublinea = Null
PrmMarca = Null
PrmClase = Null
llenaArt
End Sub

Private Sub CmdCancelar_Click()
CmdGraba.Caption = "&Nuevo"
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Mante.Visible = False
End Sub

Private Sub CmdCancelar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  CmdCancelar.Value = True
End If


End Sub


Private Sub CmdGraba_Click()
 Dim Sql2 As String
 Dim Tipo As String
 
 On Error GoTo SALE1
 Sql2 = "Insert into Tablas (tab_codcia,tab_tipreg,tab_numtab,tab_nomlargo,tab_nomcorto) values "
If CmdGraba.Caption = "&Nuevo" Then
    CmdGraba.Caption = "A&ctualizar"
    Text1.Text = ""
    Text2.Text = ""
    Text4.Text = ""
    Text2.Enabled = True
    Text4.Enabled = True
    Text2.SetFocus
    MuestraCodigo
Else
    If MsgBox("Esta Seguro de GRABAR", vbYesNo + vbQuestion, "Advertencia") = vbYes Then
          If Mante.Caption = "Mantenimiento de descuentos" Then
            Tipo = "600"
          Else
            Tipo = "333"
          End If
       Ingreso Sql2, Tipo
       CmdGraba.Caption = "&Nuevo"
       Text2.Enabled = False
       Text4.Enabled = False
       Agregalistas
    End If
End If
Exit Sub
SALE1:
MsgBox Err.Description
If Err.Number = 40002 Then Exit Sub Else MsgBox Err.Description, 48, Pub_Titulo

End Sub
Sub Ingreso(Sql2 As String, Tipo As String)
  If Tipo = 600 Then
   cnp.Execute Sql2 + "('00','" & Tipo & "','" & Val(Trim(Text1.Text)) & "','" & Trim(Text2.Text) & "','" & Trim(Text4.Text) & "')"
  Else
   cnp.Execute Sql2 + "('" & LK_CODCIA & "','" & Tipo & "','" & Val(Trim(Text1.Text)) & "','" & Trim(Text2.Text) & "','" & Trim(Text4.Text) & "')"
  End If
  
  Td.Requery
  TL.Requery
End Sub
Sub MuestraCodigo()
If Mante.Caption = "Mantenimiento de descuentos" Then
      If Td.RecordCount > 0 Then
          Td.MoveLast
          Text1.Text = Td.Fields(2) + 1
      Else
         Text1.Text = 1
      End If
Else
    If TL.RecordCount > 0 Then
       TL.MoveLast
       Text1.Text = TL.Fields(2) + 1
    Else
       Text1.Text = 1
    End If
  End If
End Sub
Private Sub CmdGraba_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  CmdCancelar.Value = True
End If
End Sub

Private Sub CmdValor_Click()
Dim dato As String
dato = (InputBox("Ingrese el valor de Llenado" + Chr(13) _
               + "Este valor se llenara en todas los" + Chr(13) _
               + "tipos de clientes y descuentos seleccionados "))
If IsNumeric(dato) Then
'If Len(dato) > 0 And Val(dato) > 0 Then
   llenarvalor dato
End If
End Sub
Sub llenarvalor(dato As String)
Dim M As String
M = "Desea Modificar a todos estos ARTICULOS por el  "
' If Val(dato) <= 0 Then Exit Sub
If MsgBox(M + dato + "%", vbYesNo) = vbYes Then
      For F = 1 To Prec.Rows - 1
           For C = 4 To Prec.Cols - 1
                 Prec.TextMatrix(F, C) = dato
            Next C
      Next F
 End If
End Sub



Private Sub Command1_Click()
If VART.SelectedItem Is Nothing Then Exit Sub
If VART.ListItems.count = 0 Then Exit Sub
  Cabecera
  llenadatos3
  buscarMostrar
End Sub

Private Sub Form_Load()
conectar
'ejecutando la creacion del comando
     set_Cmd
 'ejecutando la consulta
''''''''''''''''''''''''
     ''valores iniciales
      PrmCodCia = LK_CODCIA
      PrmTipReg = "122"
      ''''''''''''''''''''''''
Load_Familias art_familia
''''''''''''''''''''''''''
End Sub
Private Sub Form_Activate()
Compa = LK_CODCIA
F = 1
C = 4
Agregalistas
End Sub
Private Sub Cmdfiltro_Click()
If VART.ListItems.count = 0 Then Exit Sub

  Cabecera
  llenadatos2
  buscarMostrar
End Sub
Sub llenadatos2()
Dim a As Integer
Dim N As Integer
N = 0
Llenamatriz1

'''''''''''''llena lineas con descuentos
Prec.Rows = VART.ListItems.count * LstDes.SelCount + 1
Barra.max = VART.ListItems.count * LstDes.SelCount + 1
For I = 1 To VART.ListItems.count * LstDes.SelCount Step LstDes.SelCount
   N = N + 1
    Prec.TextMatrix(I, 0) = VART.ListItems(N).Text
    Prec.TextMatrix(I, 1) = VART.ListItems(N).ListSubItems(1).Text
    a = 0
    For J = 1 To LstDes.SelCount
          Prec.TextMatrix(I + a, 0) = VART.ListItems(N).Text
          Prec.TextMatrix(I + a, 2) = X(1, J)
          Prec.TextMatrix(I + a, 3) = X(2, J)
          a = a + 1
    Next J
Next I
End Sub
Sub Llenamatriz1()
        LimpiarTodo
        '''''''''''''llena matriz tipo de descuento
        Dim C As Integer
        C = 0
             With Td
                .MoveFirst
                ReDim X(2, .RecordCount)
                For I = 1 To .RecordCount
                    C = C + 1
                     If LstDes.Selected(I - 1) = True Then
                         X(1, C) = Td.Fields(2)
                         X(2, C) = Td.Fields(3)
                     Else
                         C = C - 1
                     End If
                         Td.MoveNext
                         If .EOF Then Exit For
                Next
            End With

End Sub
Sub llenadatos3()
Dim a As Integer
Dim N As Integer
Dim M As Long
Dim ACUMULA As Integer
N = 0
Llenamatriz1
For I = 1 To VART.ListItems.count
      If VART.ListItems(I).Selected = True Then
        ACUMULA = ACUMULA + 1
      End If
   Next I
   If ACUMULA = 1 Then
        TxtArt.Text = VART.SelectedItem.ListSubItems(1).Text
   Else
       TxtArt.Text = "Selecciono Varios Articulos"
   End If
   
''''''llenando matriz nueva con codigo y descripcion del articulo
     Dim Cc As Integer
        Cc = 0
                ReDim Xx(2, ACUMULA)
                For I = 1 To VART.ListItems.count
                    Cc = Cc + 1
                     If VART.ListItems(I).Selected = True Then
                         Xx(1, Cc) = VART.ListItems(I).Text
                         Xx(2, Cc) = VART.ListItems(I).ListSubItems(1).Text
                     Else
                         Cc = Cc - 1
                     End If
                       
                Next
''''''''''''''llena lineas con descuentos


Prec.Rows = ACUMULA * LstDes.SelCount + 1
Barra.max = ACUMULA * LstDes.SelCount + 1
For I = 1 To ACUMULA * LstDes.SelCount Step LstDes.SelCount
   N = N + 1
    Prec.TextMatrix(I, 0) = Xx(1, N)
    Prec.TextMatrix(I, 1) = Xx(2, N)
    a = 0
    For J = 1 To LstDes.SelCount
          Prec.TextMatrix(I + a, 0) = Xx(1, N)
          Prec.TextMatrix(I + a, 2) = X(1, J)
          Prec.TextMatrix(I + a, 3) = X(2, J)
          a = a + 1
    Next J
Next I
End Sub
Private Sub cmdgrabar_Click()
On Error GoTo SALE
If Prec.Rows = 1 Then Exit Sub
If MsgBox("Esta Seguro de GRABAR", vbYesNo + vbQuestion, "Advertencia") = vbYes Then
     grabar
End If
Exit Sub
SALE:
MsgBox Err.Description
Resume Next
If Err.Number = 40002 Then Exit Sub Else MsgBox Err.Description, 48, Pub_Titulo
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cmdsalir.Value = True
End Sub

Private Sub LstDes_GotFocus()
OSMessage1.Text = "Listado de Tipos de Descuentos / <INSERT> para agregar un nuevo rgeistro, Click Derecho para (Constante)"
End Sub

Private Sub LstDes_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 45 Then
Mante.Visible = True
Mante.Caption = "Mantenimiento de descuentos"
End If

End Sub

Private Sub LstDes_LostFocus()
OSMessage1.Text = ""
End Sub

Private Sub LstDes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 If Trim(Right$(LstDes.List(LstDes.ListIndex), 11)) = "(Constante)" Then
 
 cnp.Execute "UPDATE TABLAS SET TAB_CODCLIE='0' " & _
             "WHERE TAB_CODCIA='00' AND TAB_TIPREG='600' AND TAB_NUMTAB='" & Val(LstDes.ItemData(LstDes.ListIndex)) & "'", , adExecuteNoRecords
 Else
  cnp.Execute "UPDATE TABLAS SET TAB_CODCLIE='1' " & _
             "WHERE TAB_CODCIA='00' AND TAB_TIPREG='600' AND TAB_NUMTAB='" & Val(LstDes.ItemData(LstDes.ListIndex)) & "'", , adExecuteNoRecords
 End If
 Agregalistas
End If
End Sub

Private Sub LstLis_GotFocus()
OSMessage1.Text = "Listado de Tipos de Negocios / <INSERT> para agregar un nuevo rgeistro"
End Sub

Private Sub LstLis_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
Mante.Visible = True
Mante.Caption = "Mantenimiento de Listas"
End If
End Sub

Private Sub LstLis_LostFocus()
OSMessage1.Text = ""
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  CmdCancelar.Value = True
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text4.SetFocus
End If

End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
  VART.SetFocus
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 VART.SetFocus
End If
art_familia.ListIndex = -1
art_subfam.ListIndex = -1
art_grupo.ListIndex = -1
art_numero.ListIndex = -1
Art_Marca.ListIndex = -1
Art_Clase.ListIndex = -1

End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  CmdCancelar.Value = True
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   CmdGraba.SetFocus
End If
End Sub

Private Sub VART_DblClick()
If VART.SelectedItem Is Nothing Then Exit Sub
If VART.ListItems.count = 0 Then Exit Sub
  Cabecera
  llenadatos3
  buscarMostrar
 End Sub
Sub buscarMostrar()
Dim F As Long
Dim C As Long
Static AC As Long
Dim Sql3 As String
Barra.Visible = True
lblporcentaje.Visible = True
Barra.max = Prec.Rows - 1
AC = 0
For F = 1 To Prec.Rows - 1
    Barra.Value = AC
    AC = AC + 1
    '----------
    DoEvents
    porcentaje Prec.Rows - 1, AC
    '------------
    Compa = LK_CODCIA
    Codigo = Prec.TextMatrix(F, 0)
    VDes = Val(Trim(Left(Prec.TextMatrix(F, 2), 2)))
    '--------------------------------
    Sql3 = "select * from clidscto where cld_codcia='" & LK_CODCIA & "' and cld_codart=" & Codigo & " and CLD_tipodscto=" & VDes
    DS.Close
    DS.Open Sql3, cnp, adOpenForwardOnly, adLockOptimistic, adCmdText
      For C = 4 To Prec.Cols - 1
            VLis = Val(Trim(Left(Prec.TextMatrix(0, C), 2)))
            Prec.TextMatrix(F, C) = buscar(Compa, Codigo, VDes, VLis)
      Next C
Next F
Barra.Visible = False
lblporcentaje.Visible = False
End Sub
Function buscar(comp As String, Codigo As String, VDes As Integer, VLis As Integer) As Single
Dim F As Long
Dim C As Long
   With DS
       If .RecordCount > 0 Then
         .MoveFirst
      End If
      '''''''''''''recorre la tabla
        For F = 1 To .RecordCount
            If !cld_codcia = comp And !CLD_CODART = Val(Codigo) And Trim(!CLD_tipodscto) = VDes And Trim(!CLD_listadscto) = VLis Then
               buscar = !cld_desto1
               Exit Function
            Else
              buscar = 0
            End If
            .MoveNext
            If .EOF Then Exit Function
        Next
    End With
End Function
Sub grabar()
Dim E As Long
Dim N As Integer
Dim SQL As String
Static Ad As Long
Barra.Visible = True
lblporcentaje.Visible = True
Ad = 0
Barra.max = Prec.Rows - 1
SQL = ""
 SQL = "insert into clidscto values "
 With DS
      For F = 1 To Prec.Rows - 1
          Barra.Value = Ad
          Ad = Ad + 1
          DoEvents
          porcentaje Prec.Rows - 1, Ad
          
           For C = 4 To Prec.Cols - 1
               Codigo = Prec.TextMatrix(F, 0)
               VDes = Val(Trim(Left(Prec.TextMatrix(F, 2), 2)))
               VLis = Val(Trim(Left(Prec.TextMatrix(0, C), 2)))
               VALOR = Val(Prec.TextMatrix(F, C))
               '''''''''''''''''''''
               cnp.Execute "delete from clidscto where cld_codcia='" & LK_CODCIA & "' and cld_codart='" & Codigo & "' and CLD_tipodscto='" & VDes & "' and CLD_listadscto='" & VLis & "'"
               DS.Requery
               ''''''''''''''''''''
                  If VALOR > 0 Then
                      cnp.Execute SQL + "('" & LK_CODCIA & "','" & VDes & "','" & Codigo & "','" & VLis & "','" & VALOR & "')"
               '''''''''''''''''''''
                  End If
          Next C
   Next F
   
  End With
 DS.Requery
 Barra.Visible = False
 lblporcentaje.Visible = False
End Sub
Sub Agregalistas()
LstLis.Clear
LstDes.Clear
TL.Requery
Td.Requery
'''''''''''''''LLENA LISTA
With TL
     .MoveFirst
     For I = 1 To .RecordCount
       If !TAB_TIPREG = "333" Then
           LstLis.AddItem Trim(!tab_NOMLARGO)
           LstLis.Selected(I - 1) = True
       End If
         .MoveNext
          If .EOF Then Exit For
     Next I
'''''''''''''''LLENA LISTA
  End With
     With Td
     .MoveFirst
     For I = 1 To .RecordCount
       If !TAB_TIPREG = "600" Then
           If !TAB_CODCLIE = 1 Then
            LstDes.AddItem Trim(!tab_NOMLARGO) + "   " + "(Constante)"
           Else
            LstDes.AddItem Trim(!tab_NOMLARGO)
           End If
           LstDes.ItemData(LstDes.NewIndex) = Val(!TAB_NUMTAB)
           LstDes.Selected(I - 1) = True
             .MoveNext
          If .EOF Then Exit For
       End If
     Next I
       
End With
Cabecera
End Sub
Private Sub Text3_Change()
Dim SQL As String
SQL = "select * from Arti where art_codcia='" & LK_CODCIA & "' and art_nombre like '" & Trim(Text3.Text) & "%'"
With Ar
      .CursorLocation = adUseClient
      .Open SQL, cnp, adOpenKeyset, adLockOptimistic
      .Requery
      VART.ListItems.Clear
      While Not .EOF
         Clave = Ar!art_KEY & "A"
         Set mItem = VART.ListItems.Add(Key:=Clave, Text:=Ar!art_KEY)
         mItem.ListSubItems.Add Text:=Ar!art_nombre
         mItem.ListSubItems.Add Text:=Ar!art_KEY
        .MoveNext
      Wend
  End With
Ar.Close
End Sub
Sub Cabecera()
 Dim Cabecera As String
 '''llenando caja de Texto
   If numart = 1 Then
    Prec.Rows = LstDes.SelCount + 1
    Prec.TextMatrix(1, 1) = VART.SelectedItem.ListSubItems(1).Text
  End If
''TITULO DE LOS 3 PRIMERAS COLUMNAS
  
  Cabecera = "Codigo|<Articulo                           |<Codigo|<Tipo Lista Dscto     "
  vueltas = 0
  TL.Requery
  TL.MoveFirst
  Prec.Cols = 4
  With TL
      For I = 1 To LstLis.ListCount
       If LstLis.Selected(I - 1) = True Then
           Cabecera = Cabecera + "|<" + CStr(.Fields(2)) + "  " + .Fields(4) + " "
       End If
             .MoveNext
             If .EOF Then Exit For
        Next
         Prec.Cols = LstLis.SelCount + 4
        .MoveFirst
        ''''''''''''''''''''''
        For I = 4 To Prec.Cols
             ''configuracion de columnas
                Prec.ColEdit(I) = True
                Prec.ColControl(I) = osTextBox
                Prec.ColData(I) = osDecimal
                Prec.ColSalto(I) = 1
                  If I > 2 Then
                   Prec.ColAtras(I) = 1
                  End If
        Next
          Prec.FormatString = Cabecera
           Prec.ColWidth(0) = 0
          Prec.ColWidth(2) = 0
  End With
End Sub

Sub LimpiarTodo()
    For F = 1 To Prec.Rows - 1
       For C = 1 To Prec.Cols - 1
            Prec.TextMatrix(F, C) = ""
       Next C
   Next F
 End Sub
Private Sub Cmdsalir_Click()
Set Ta = Nothing
Set Pr = Nothing
 Set Ar = Nothing
Set DS = Nothing
Set Td = Nothing
Set cnp = Nothing
Set TL = Nothing
'Set cnp = Nothing
Set CmdFamilias = Nothing
Set CmdIp = Nothing
Set PrmCodArt = Nothing
Set PrmTipReg = Nothing
Set PrmCodArt = Nothing
Set PrmCodCia = Nothing
Set PrmPack = Nothing
Set PrmDiv = Nothing
Set PrmFam = Nothing
Set PrmLinea = Nothing
Set PrmSublinea = Nothing
Set PrmMarca = Nothing
Set PrmClase = Nothing

Set PrmDesc = Nothing
Set PrmAlterno = Nothing
Unload Me
End Sub
Sub conectar()
'Estableciendo la Coneccion
Dim cnpt As String
    cnpt = "Provider=MSDataShape;data Provider=MSDASQL;Driver={SQL Server};"
         cnpt = cnpt + "Server=" & CONST_SERVER & ""
         cnpt = cnpt + ";uid=" & CONST_UID & ""
         cnpt = cnpt + ";pwd=" & CONST_PWD & ""
         cnpt = cnpt + ";database=BDATOS"
    cnp.ConnectionTimeout = 15
    cnp.Open cnpt
'''''''''''''''''''''
 With TL
      .CursorLocation = adUseClient
      .Open "select * from Tablas where tab_codcia='" & LK_CODCIA & "' and tab_tipreg='333'", cnp, adOpenKeyset, adLockOptimistic
      .Requery
      
  End With
 With Td
      .CursorLocation = adUseClient
      .Open "select * from Tablas where tab_codcia='00' and tab_tipreg='600'", cnp, adOpenKeyset, adLockOptimistic
      .Requery
  End With
  With DS
       .CursorLocation = adUseClient
       .Open "select * from clidscto", cnp, adOpenForwardOnly, adLockOptimistic, adCmdText
       .Requery
  End With
  
End Sub
Private Sub VART_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If VART.SelectedItem Is Nothing Then Exit Sub
        If VART.ListItems.count = 0 Then Exit Sub
          Cabecera
          llenadatos3
          buscarMostrar
End If
End Sub

Private Sub VART_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    VART.ToolTipText = "Existen  " + str(VART.ListItems.count) + "   Articulos"
End Sub
Sub porcentaje(NUMERO As Long, LINEAS As Long)
Dim cad As String * 4
If NUMERO = LINEAS Then
  cad = "%"
Else
  cad = "%..."
End If
lblporcentaje.Caption = Format((LINEAS * 100) / NUMERO, "0") & cad
End Sub

