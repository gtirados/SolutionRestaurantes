VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmPreUnidad 
   Caption         =   "Actualización de Precios / Stock Min.-Max."
   ClientHeight    =   6255
   ClientLeft      =   210
   ClientTop       =   885
   ClientWidth     =   8430
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "FrmPreUnidad.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   8430
   WindowState     =   2  'Maximized
   Begin VB.Frame fracambia 
      Caption         =   "Actualizacion de Precios"
      Height          =   375
      Left            =   8160
      TabIndex        =   30
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
      Begin MSFlexGridLib.MSFlexGrid gridp 
         Height          =   2055
         Left            =   4440
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3625
         _Version        =   393216
         FocusRect       =   2
         HighLight       =   0
         AllowUserResizing=   1
      End
      Begin VB.CommandButton preliminar 
         Caption         =   "&Preliminar"
         Height          =   375
         Left            =   5280
         TabIndex        =   48
         Top             =   3000
         Width           =   2895
      End
      Begin VB.CommandButton propre 
         Caption         =   "Cambiar Lista de Precios"
         Height          =   375
         Left            =   1920
         TabIndex        =   47
         Top             =   3720
         Width           =   5535
      End
      Begin VB.TextBox pre5 
         Height          =   285
         Left            =   4080
         MaxLength       =   10
         TabIndex        =   45
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox pre4 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   43
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox pre3 
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   41
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox pre2 
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   39
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox pre1 
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   37
         Top             =   3120
         Width           =   855
      End
      Begin VB.ListBox grupo 
         Height          =   1860
         Left            =   3120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.ListBox subfami 
         Height          =   1185
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   32
         Top             =   1320
         Width           =   2895
      End
      Begin VB.ComboBox fami 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label eti 
         Caption         =   "Lista de Productos"
         Height          =   255
         Index           =   8
         Left            =   4680
         TabIndex        =   49
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label eti 
         Caption         =   "Precio 5"
         Height          =   255
         Index           =   7
         Left            =   4080
         TabIndex        =   46
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label eti 
         Caption         =   "Precio 4"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   44
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label eti 
         Caption         =   "Precio 3"
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   42
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label eti 
         Caption         =   "P.Minor."
         Height          =   255
         Index           =   4
         Left            =   1200
         TabIndex        =   40
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label eti 
         Caption         =   "P.Mayor."
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label eti 
         Caption         =   "Grupo"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label eti 
         Caption         =   "Sub Familia"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label eti 
         Caption         =   "Familia"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   2040
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ListBox ListUnidad 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   4800
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton cmdcompras 
      Caption         =   "   ===> &Ultimas Compras"
      Height          =   735
      Left            =   0
      TabIndex        =   21
      Top             =   5400
      Width           =   730
   End
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "&Grabar Datos"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   18
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox textovar 
      Height          =   285
      Left            =   5040
      MaxLength       =   9
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid Grid_Arti 
      Height          =   3135
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5530
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid grid_det 
      Height          =   1455
      Left            =   0
      TabIndex        =   15
      Top             =   3705
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   2566
      _Version        =   393216
      Rows            =   3
      Cols            =   5
      FixedRows       =   2
      FixedCols       =   4
      Enabled         =   0   'False
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   9
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame fra2 
      Height          =   3255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtdato 
         Height          =   315
         Left            =   120
         TabIndex        =   52
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton CmdProcesa 
         Caption         =   "&Listar  Articulo (s)"
         Height          =   495
         Left            =   1440
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Proveedor"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   12
         Top             =   1005
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Familia"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Busqueda por Descripcion,Marca y Color"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.ListBox PROV 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   1800
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ListBox lineas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ListBox listapro 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   4
         Top             =   2310
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_key 
         Height          =   315
         Left            =   240
         MaxLength       =   8
         TabIndex        =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Busqueda por Descripcion,Marca,Color:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub Familia"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Familia :"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   910
         Width           =   570
      End
      Begin VB.Label lblopcion 
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2445
      End
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grid_compras 
      Height          =   2775
      Left            =   900
      TabIndex        =   19
      Top             =   5160
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   6
      Cols            =   5
      FocusRect       =   2
      HighLight       =   2
      AllowUserResizing=   1
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECIO. 6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   9960
      TabIndex        =   51
      Top             =   3405
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LBLARTI 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label lblcosto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COSTO"
      Height          =   315
      Left            =   2400
      TabIndex        =   28
      Top             =   3405
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECIO. 5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   8880
      TabIndex        =   27
      Top             =   3405
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECIO. 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   7440
      TabIndex        =   26
      Top             =   3405
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECIO. 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   6000
      TabIndex        =   25
      Top             =   3405
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECIO. 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   4560
      TabIndex        =   24
      Top             =   3405
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PRECIO. 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3120
      TabIndex        =   23
      Top             =   3405
      Width           =   1400
   End
   Begin VB.Label lblmensa 
      Caption         =   "Procesando  . . . "
      Height          =   255
      Left            =   3960
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "FrmPreUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loc_key  As Integer
Dim PSART_LOC As rdoQuery
Dim PSART_LOC2 As rdoQuery
Dim PSPRO_V As rdoQuery
Dim artloc_llave As rdoResultset
Dim artloc_mayor As rdoResultset
Dim PRO_V As rdoResultset
Dim PRE_UNIDADES As rdoQuery
Dim preuni_llave As rdoResultset
Dim PRE_UNIDADES2 As rdoQuery
Dim preuni_llave2 As rdoResultset
Dim Flag_Consis As String * 1
Dim Flag_F2 As String * 1
Dim Flag_Bloq As String * 1
Dim Flag_Inicial As String * 1
Dim Flag_Change  As String * 1
Dim loc_fila  As Integer
Dim loc_colum  As Integer
Dim loc_unid As String * 1
Dim WCOD_ORIGINAL As Currency
Dim loc_flag_stock As String * 1
Dim loc_flag_por As String * 1
Dim loc_flag_precios As String * 1
Dim loc_comis(5)  As String * 10
Dim VAR_ACTIVAR As Integer
Dim wnormal As Integer
Dim LOC_EQUIV As Currency


Private Sub PROCESAR_DATOS()
Dim Modo As String * 1
Dim clave_far As String
Dim suma_cantid As Currency
Dim ENTRADA
Dim WALT As String * 20
lblmensa.Visible = True
Modo = ""
If option1(0).Value Then
   If val(Txt_key.Text) <= 0 And LK_FLAG_ALTERNO <> "A" Then
     lblmensa.Visible = False
     MsgBox "Seleccione Codigo, para Procesar ", 48, Pub_Titulo
     GoTo fin
   End If
   If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
     ENTRADA = WCOD_ORIGINAL
   Else
     ENTRADA = val(Txt_key.Text)
   End If
   pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_KEY = ?  AND ART_CALIDAD = 1 ORDER BY ART_ALTERNO"
   Modo = "1"
ElseIf option1(1).Value Then
  If lineas.ListCount = 0 And listapro.ListCount = 0 Then
    lblmensa.Visible = False
    MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
    GoTo fin
  End If
  If Trim(lineas.Text) = "" And Trim(listapro.Text) = "" Then
    lblmensa.Visible = False
    MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
    GoTo fin
  End If
  If Trim(listapro.Text) = "" And Trim(lineas.Text) <> "" Then
     pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_FAMILIA = ? AND ART_CALIDAD = 1 ORDER BY ART_ALTERNO"
  ElseIf Trim(listapro.Text) <> "" And Trim(lineas.Text) = "" Then
     pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_SUBFAM = ? AND ART_CALIDAD = 1 ORDER BY ART_ALTERNO"
  Else
    pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_FAMILIA = ? AND ART_SUBFAM = ? AND ART_CALIDAD = 1 ORDER BY ART_ALTERNO"
  End If
  Modo = "2"
ElseIf option1(2).Value Then
   If Trim(PROV.Text) = "" Then
     lblmensa.Visible = False
     MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
     GoTo fin
   End If
   Wche = 0
   wkSELECT = ""
   For i = 0 To PROV.ListCount - 1
     PROV.ListIndex = i
     If PROV.Selected(i) Then
       Wche = 1
       wkSELECT = wkSELECT + " ART_CODCLIE = " & Trim(Right(PROV.Text, 10)) & " OR"
      End If
   Next i
   If Wche = 0 Then
     lblmensa.Visible = False
     MsgBox "Seleccione Datos, para Procesar ", 48, Pub_Titulo
     GoTo fin
   End If
   wkSELECT = Left(wkSELECT, Len(wkSELECT) - 3)
   pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ? AND " & wkSELECT & " AND ART_CALIDAD = 1 ORDER BY ART_ALTERNO"
   Modo = "3"
End If

Set PSART_LOC = CN.CreateQuery("", pub_cadena)
If Modo = "1" Then
   PSART_LOC(0) = 0
   PSART_LOC(1) = 0
ElseIf Modo = "2" Then
    If Trim(listapro.Text) = "" And Trim(lineas.Text) <> "" Then
       PSART_LOC(0) = 0
       PSART_LOC(1) = 0
    ElseIf Trim(listapro.Text) <> "" And Trim(lineas.Text) = "" Then
       PSART_LOC(0) = 0
       PSART_LOC(1) = 0
    Else
       PSART_LOC(0) = 0
       PSART_LOC(1) = 0
       PSART_LOC(2) = 0
    End If
ElseIf Modo = "3" Then
   PSART_LOC(0) = ws_codcia
End If
 

Set artloc_llave = PSART_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
ws_codcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
  ws_codcia = "00"
End If

If Modo = "1" Then
   PSART_LOC(0) = ws_codcia
   PSART_LOC(1) = ENTRADA
ElseIf Modo = "2" Then
    If Trim(listapro.Text) = "" And Trim(lineas.Text) <> "" Then
       PSART_LOC(0) = ws_codcia
       PSART_LOC(1) = Right(lineas.Text, 10)
    ElseIf Trim(listapro.Text) <> "" And Trim(lineas.Text) = "" Then
       PSART_LOC(0) = ws_codcia
       PSART_LOC(1) = Right(listapro.Text, 10)
    Else
       PSART_LOC(0) = ws_codcia
       PSART_LOC(1) = Right(lineas.Text, 10)
       PSART_LOC(2) = Right(listapro.Text, 10)
    End If
ElseIf Modo = "3" Then
   PSART_LOC(0) = ws_codcia
End If
  artloc_llave.Requery
  If artloc_llave.EOF Then
    lblmensa.Visible = False
    PB.Visible = False
    MsgBox "No existen Articulos en la Lista", 48, Pub_Titulo
    GoTo fin
  End If
  lblmensa.Caption = "Procesando . . . "
  DoEvents
  PB.max = artloc_llave.RowCount
  PB.Min = 0
  PB.Value = 0
  PB.Visible = True
  DoEvents
  Fila = 1
  Flag_Inicial = "A"
  Grid_Arti.Clear
  Grid_Arti.Rows = 3
  Grid_Arti.FixedRows = 2
  CABEZA_ARTI
  PRE_UNIDADES(0) = LK_CODCIA
  PRE_UNIDADES(2) = "A"
  WALT = String(8, " ")
  pu_codcia = LK_CODCIA
  SQ_OPER = 1
  Do Until artloc_llave.EOF
    PB.Value = PB.Value + 1
    PRE_UNIDADES(1) = val(artloc_llave!ART_KEY)
    preuni_llave.Requery
    PUB_CODART = val(artloc_llave!ART_KEY)
    LEER_ARM_LLAVE
    Fila = Fila + 1
    Grid_Arti.Rows = Fila + 1
    Grid_Arti.Row = Fila
    Grid_Arti.RowHeight(Fila) = 285
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      WALT = Trim(artloc_llave!ART_ALTERNO)
      Grid_Arti.TextMatrix(Fila, 0) = "'" & WALT + "-" + Trim(artloc_llave!art_nombre)
    Else
      WALT = Trim(artloc_llave!ART_KEY)
      Grid_Arti.TextMatrix(Fila, 0) = "'" & WALT + "-" + Trim(artloc_llave!art_nombre)
    End If
    lblmensa.Caption = "Procesando . . .  " & Trim(artloc_llave!art_nombre)
    DoEvents
    If preuni_llave.EOF Then
     GoTo NOMUESTRA
    End If
    wsCostoRep = Nulo_Valor0(preuni_llave!PRE_COSTO) ' COSTO_REPOS(artloc_llave!art_KEY)
    Grid_Arti.TextMatrix(Fila, 1) = preuni_llave!pre_unidad
    Grid_Arti.TextMatrix(Fila, 2) = Format(arm_llave!arm_cospro * preuni_llave!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Fila, 3) = Format(wsCostoRep * preuni_llave!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Fila, 4) = Format(preuni_llave!PRE_COSTO * 1.19, "0.0000")
    'If LK_EMP <> "PIU" Then
    Grid_Arti.TextMatrix(Fila, 5) = preuni_llave!PRE_PRE1
    Grid_Arti.TextMatrix(Fila, 6) = preuni_llave!PRE_PRE2
   ' Else
   '   Grid_Arti.TextMatrix(fila, 5) = Nulo_Valors(preuni_llave!PRE_PRE1)
   '   Grid_Arti.TextMatrix(fila, 6) = Nulo_Valors(preuni_llave!PRE_PRE2)
   ' End If
    
    Grid_Arti.TextMatrix(Fila, 7) = artloc_llave!ART_KEY
    Grid_Arti.TextMatrix(Fila, 8) = preuni_llave!PRE_EQUIV
    Grid_Arti.TextMatrix(Fila, 9) = Nulo_Valor0(arm_llave!arm_cospro)
    Grid_Arti.TextMatrix(Fila, 10) = wsCostoRep
    Grid_Arti.TextMatrix(Fila, 11) = Format(Nulo_Valor0(artloc_llave!art_stock_min) / preuni_llave!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Fila, 12) = Format(Nulo_Valor0(artloc_llave!ART_STOCK_MAX) / preuni_llave!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Fila, 13) = Nulo_Valor0(artloc_llave!ART_POR1)
    Grid_Arti.TextMatrix(Fila, 14) = Nulo_Valor0(artloc_llave!ART_POR2)
    Grid_Arti.TextMatrix(Fila, 15) = Nulo_Valor0(artloc_llave!ART_POR3)
    Grid_Arti.TextMatrix(Fila, 16) = Nulo_Valor0(artloc_llave!ART_POR4)
    Grid_Arti.TextMatrix(Fila, 17) = Nulo_Valor0(artloc_llave!ART_POR5)
    Grid_Arti.TextMatrix(Fila, 19) = Nulo_Valor0(artloc_llave!ART_STOCK_MAX)
    Grid_Arti.TextMatrix(Fila, 20) = Nulo_Valor0(artloc_llave!art_stock_min)
    Grid_Arti.COL = 18
    Grid_Arti.CellForeColor = QBColor(2)
 

NOMUESTRA:
   artloc_llave.MoveNext
  Loop
  PB.Visible = False
  lblmensa.Visible = False
  Grid_Arti.COL = 1
  Grid_Arti.Row = 2
  Flag_Inicial = ""
'  Grid_Arti.SetFocus
Exit Sub

GoTo fin
Exit Sub
fin:
fra2.Enabled = True
PB.Visible = False
lblmensa.Visible = False
DoEvents
CmdProcesa.Enabled = True
DoEvents
'opciones.Enabled = True

End Sub


Private Sub cmdcompras_Click()
Dim Mensaje, Título, valorpred, mivalor
    Mensaje = "Cuantas ultimos compras desea ver : "
    Título = "Ultmas Compras "
    valorpred = "5"
    mivalor = InputBox(Mensaje, Título, valorpred)
    If mivalor = "" Or val(mivalor) = 0 Then
       Exit Sub
    End If
    If Len(Trim(mivalor)) > 4 Or Len(Trim(mivalor)) = 0 Then
       MsgBox "Demasiados Datos", 48, Pub_Titulo
       Exit Sub
    End If
    If InStr(1, mivalor, ".") > 0 Then
       MsgBox "Debe ser Entero  .... ", 48, Pub_Titulo
       Exit Sub
    End If
    If Not IsNumeric(mivalor) Then
       MsgBox "Debe ser Numerico .... ", 48, Pub_Titulo
       Exit Sub
    End If
    Screen.MousePointer = 11
    cmdcompras.Enabled = False
    COMPRAS_ULT val(mivalor)
    cmdcompras.Enabled = True
    grid_det.SetFocus
    Screen.MousePointer = 0
End Sub

Private Sub cmdgrabar_Click()
On Error GoTo mensa
If loc_flag_precios <> "A" Then
 MsgBox "No Tiene Ascceso ... ", 48, Pub_Titulo
 Exit Sub
End If
If TEXTOVAR.Visible = True Then
   MsgBox "Esta Editando Datos, No procede...", 48, Pub_Titulo
   Exit Sub
End If
If grid_det.Enabled = False Then
   Exit Sub
End If
If Not CONSIS_UNIDAD Then
    Screen.MousePointer = 0
    grid_det.SetFocus
    Exit Sub
End If

pub_mensaje = " Grabar todas la modificaciones.  ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   grid_det.SetFocus
   Exit Sub
End If
preuni_llave2.MoveFirst
If preuni_llave2.EOF Then
 MsgBox " Llave Fuera preuni_llave2...", 48, Pub_Titulo
 Exit Sub
End If
Fila = 0
Flag_Inicial = "A"
Do Until preuni_llave2.EOF
   Fila = Fila + 1
   preuni_llave2.Edit
   preuni_llave2!PRE_COSTO = val(grid_det.TextMatrix(Fila, 2)) / 1.19
   preuni_llave2!PRE_POR1 = val(grid_det.TextMatrix(Fila, 4))
   preuni_llave2!PRE_PRE1 = val(grid_det.TextMatrix(Fila, 5))
   preuni_llave2!PRE_POR2 = val(grid_det.TextMatrix(Fila, 6))
   preuni_llave2!PRE_PRE2 = val(grid_det.TextMatrix(Fila, 7))
   preuni_llave2!PRE_POR3 = val(grid_det.TextMatrix(Fila, 8))
   preuni_llave2!PRE_PRE3 = val(grid_det.TextMatrix(Fila, 9))
   preuni_llave2!PRE_POR4 = val(grid_det.TextMatrix(Fila, 10))
   preuni_llave2!PRE_PRE4 = val(grid_det.TextMatrix(Fila, 11))
   preuni_llave2!PRE_POR5 = val(grid_det.TextMatrix(Fila, 12))
   preuni_llave2!PRE_PRE5 = val(grid_det.TextMatrix(Fila, 13))
   preuni_llave2!PRE_COSTO_REPO = val(grid_det.TextMatrix(Fila, 14))
   preuni_llave2.Update
   If preuni_llave2!PRE_FLAG_UNIDAD = "A" Then
    Grid_Arti.TextMatrix(Grid_Arti.Row, 2) = Format(val(Grid_Arti.TextMatrix(Grid_Arti.Row, 9)) * preuni_llave2!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Grid_Arti.Row, 3) = Format(val(Grid_Arti.TextMatrix(Grid_Arti.Row, 10)) * preuni_llave2!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Grid_Arti.Row, 4) = Format(preuni_llave2!PRE_COSTO, "0.0000")
    Grid_Arti.TextMatrix(Grid_Arti.Row, 5) = preuni_llave2!PRE_COSTO_REPO
    Grid_Arti.TextMatrix(Grid_Arti.Row, 6) = preuni_llave2!PRE_PRE1
    Grid_Arti.TextMatrix(Grid_Arti.Row, 8) = preuni_llave2!PRE_EQUIV
   End If
   preuni_llave2.MoveNext
Loop
'MsgBox "Ok. Actualizado .", 48, Pub_Titulo
LBLARTI.Caption = ""
Flag_Inicial = ""
fra2.Enabled = True
grid_det.Clear
grid_det.Cols = 1
grid_det.Rows = 1
grid_det.Enabled = False
Grid_Arti.Row = loc_fila
For Fila = -6 To -1
  Grid_Arti.COL = Fila * -1
  Grid_Arti.CellForeColor = QBColor(0)
  Grid_Arti.CellBackColor = QBColor(15)
Next Fila
cmdGrabar.Enabled = False
grid_compras.Clear
grid_compras.Enabled = False
cmdcompras.Enabled = False
'FrmPreUnidad
Grid_Arti.SetFocus
Exit Sub
mensa:
If Err.Number = 40002 Then
 MsgBox "Verificar Datos de Porcentajes, deben ser menores de 100% - Intente Nuevamente ", 48, Pub_Titulo
Else
 MsgBox "Verificar Datos , Intente Nuevamente ", 48, Pub_Titulo
End If
End Sub

Private Sub cmdlista_Click()
VAR_ACTIVAR = 0
 fra2.Enabled = False
 DoEvents
 Grid_Arti.Visible = False
 SALIR.Enabled = False
 Screen.MousePointer = 11
 PROCESAR_DATA slDescripcion, slMarca, slMedida, slCodigo, sIndex
 Screen.MousePointer = 0
 SALIR.Enabled = True
 If Grid_Arti.Rows < 2 Then
    Exit Sub
 End If
 
 If Grid_Arti.Rows = 3 Then
  Grid_Arti.TextMatrix(1, 0) = Left(Grid_Arti.TextMatrix(1, 0), 8) & " (" & Grid_Arti.Rows - 2 & ")"
 Else
  Grid_Arti.TextMatrix(1, 0) = Grid_Arti.TextMatrix(1, 0) & "(" & Grid_Arti.Rows - 2 & ")"
 End If
 fra2.Enabled = True
 grid_det.Enabled = False
 grid_compras.Enabled = False
 cmdGrabar.Enabled = False
 Grid_Arti.Visible = True
 Grid_Arti.Enabled = True
 Grid_Arti.SetFocus
End Sub

Private Sub CmdProcesa_Click()
 VAR_ACTIVAR = 0
 fra2.Enabled = False
 DoEvents
 Grid_Arti.Visible = False
 SALIR.Enabled = False
 Screen.MousePointer = 11
 PROCESAR_DATOS
 Screen.MousePointer = 0
 SALIR.Enabled = True
 If Grid_Arti.Rows < 2 Then
    Exit Sub
 End If
 
 If Grid_Arti.Rows = 3 Then
  Grid_Arti.TextMatrix(1, 0) = Left(Grid_Arti.TextMatrix(1, 0), 8) & " (" & Grid_Arti.Rows - 2 & ")"
 Else
  Grid_Arti.TextMatrix(1, 0) = Grid_Arti.TextMatrix(1, 0) & "(" & Grid_Arti.Rows - 2 & ")"
 End If
 fra2.Enabled = True
 grid_det.Enabled = False
 grid_compras.Enabled = False
 cmdGrabar.Enabled = False
 Grid_Arti.Visible = True
 Grid_Arti.Enabled = True
 Grid_Arti.SetFocus
End Sub

Private Sub CmdProcesa_GotFocus()
If loc_unid = "A" Then
 Exit Sub
End If
If Grid_Arti.Enabled = False Then
 Exit Sub
End If
grid_det.Cols = 1
grid_det.Rows = 1
cmdGrabar.Enabled = False
grid_det.Clear
grid_det.Enabled = False
Grid_Arti.Clear
Grid_Arti.Rows = 1
Grid_Arti.Cols = 1
Grid_Arti.Enabled = False
grid_compras.Rows = 3
grid_compras.Enabled = False
cmdcompras.Enabled = False

End Sub

Private Sub fami_Click()
Dim wpos As Integer
Dim WFAMI2 As Integer
If Trim(fami.Text) = "" Then
 subfami.Clear
 Exit Sub
End If
Screen.MousePointer = 11
wpos = subfami.ListIndex
WFAMI2 = val(Trim(Right(fami.Text, 6)))
FAMI_SUBFAM WFAMI2
On Error GoTo sigue
subfami.ListIndex = wpos
Screen.MousePointer = 0
Exit Sub
sigue:
Resume Next

End Sub

Private Sub Form_Load()
Dim fech As String
Dim wmes As String * 10
Dim VARX As String
Dim r As Integer
Dim WT As Integer
wnormal = 1 ' Siempre calida Normal es codigo 1
loc_flag_por = " "
loc_flag_stock = " "
loc_flag_precios = " "
For Fila = 1 To lk_OTROS_Count
   If val(lk_OTROS(Fila)) = 3 Then '% Comision
    loc_flag_por = "A"
   End If
   If val(lk_OTROS(Fila)) = 2 Then ' Stock Min-Max
    loc_flag_stock = "A"
   End If
   If val(lk_OTROS(Fila)) = 1 Then ' Precios
    loc_flag_precios = "A"
   End If
Next Fila
Flag_Change = "A"
Flag_Inicial = ""
pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? AND PRE_FLAG_UNIDAD = ? ORDER BY PRE_CODART"
Set PRE_UNIDADES = CN.CreateQuery("", pub_cadena)
PRE_UNIDADES(0) = 0
PRE_UNIDADES(1) = 0
PRE_UNIDADES(2) = 0
Set preuni_llave = PRE_UNIDADES.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_CODART, PRE_SECUENCIA"
Set PRE_UNIDADES2 = CN.CreateQuery("", pub_cadena)
PRE_UNIDADES2(0) = 0
PRE_UNIDADES2(1) = 0
Set preuni_llave2 = PRE_UNIDADES2.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP = 'P'  AND CLI_CODCIA = ? ORDER BY CLI_NOMBRE"
Set PSPRO_V = CN.CreateQuery("", pub_cadena)
PSPRO_V(0) = 0
Set PRO_V = PSPRO_V.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSPRO_V(0) = LK_CODCIA
PRO_V.Requery
Do Until PRO_V.EOF
    PROV.AddItem PRO_V!CLI_NOMBRE & String(15, " ") & PRO_V!cli_codclie
    PRO_V.MoveNext
Loop
PROV.Visible = False
PUB_CODCIA = LK_CODCIA
If LK_EMP_PTO = "A" Then
    PUB_CODCIA = "00"
End If
LLENADOS lineas, 122
LLENADOS listapro, 123
ListView1.Visible = False
Txt_key.Visible = False
lineas.Visible = False
listapro.Visible = False
PROV.Visible = False
Label1.Visible = False
txtdato.Visible = True
Label1.Caption = ""
Label2.Visible = False
CABEZA_ARTI
CABEZA_DET
Txt_key.MaxLength = 15
SQ_OPER = 2
PUB_TIPREG = 45
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
  Label3(tab_mayor!TAB_NUMTAB - 1).Caption = Trim(tab_mayor!tab_NOMLARGO)
 'loc_comis(tab_mayor!TAB_NUMTAB) = Trim(tab_mayor!tab_NOMLARGO)
  tab_mayor.MoveNext
Loop

If LK_EMP = "HER" Then
  fracambia.Visible = False
'  txtdato.SetFocus
  LLENA_COMBO fami, 122
  LLENADOS grupo, 123
End If

End Sub

Private Sub Grid_Arti_DblClick()
If ListUnidad.Visible Then
 ListUnidad.Visible = False
 loc_unid = ""
 Exit Sub
End If
Grid_Arti_KeyPress 13
End Sub

Private Sub Grid_Arti_EnterCell()
If Flag_Inicial = "A" Then
 Exit Sub
End If
If LK_EMP = "PIU" And Grid_Arti.COL >= 5 Then
  ElGrid_EnterCell Grid_Arti, TEXTOVAR
End If
If Grid_Arti.COL >= 11 Then
 ElGrid_EnterCell Grid_Arti, TEXTOVAR
End If


End Sub

Private Sub Grid_Arti_GotFocus()
If ListUnidad.Visible Then
 ListUnidad.Visible = False
 loc_unid = ""
 Exit Sub
End If
If loc_unid = "A" Then
 loc_unid = ""
 Exit Sub
End If
 If grid_det.Enabled Then
   grid_det.Cols = 1
   grid_det.Rows = 1
   grid_det.Clear
   cmdGrabar.Enabled = False
   grid_det.Enabled = False
   cmdGrabar.Enabled = False
   Grid_Arti.Row = loc_fila
   For Fila = -6 To -1
     Grid_Arti.COL = Fila * -1
     Grid_Arti.CellForeColor = QBColor(0)
     Grid_Arti.CellBackColor = QBColor(15)
   Next Fila
   grid_compras.Clear
   grid_compras.Enabled = False
   cmdcompras.Enabled = False
   Grid_Arti.SetFocus
 End If

End Sub


Private Sub Grid_Arti_KeyDown(KeyCode As Integer, Shift As Integer)
Exit Sub
If KeyCode = 32 Then
  If Grid_Arti.COL = 1 Then
    If Trim(Grid_Arti.TextMatrix(Grid_Arti.Row, 1)) = "" Then
       Exit Sub
    End If
    Dim wsUnidad As String * 15
    Dim wsEquiv As String * 5
    'pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ? ORDER BY PRE_CODART, PRE_SECUENCIA"
    'Set PRE_UNIDADES2 = CN.CreateQuery("", pub_cadena)
    'Set preuni_llave2 = PRE_UNIDADES2.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
    PRE_UNIDADES2(0) = LK_CODCIA
    PRE_UNIDADES2(1) = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 7))
    preuni_llave2.Requery
    If preuni_llave2.EOF Then
     Exit Sub
    End If
    ListUnidad.Clear
    ListUnidad.Visible = True
    Do Until preuni_llave2.EOF
     wsUnidad = Trim(preuni_llave2!pre_unidad)
     wsEquiv = str(preuni_llave2!PRE_EQUIV)
     ListUnidad.AddItem wsUnidad & "   " & wsEquiv & String(40, " ") & preuni_llave2!pre_secuencia
     preuni_llave2.MoveNext
    Loop
    loc_unid = "A"
    ListUnidad.ListIndex = 0
    ListUnidad.SetFocus
  End If
 Exit Sub
End If
If Flag_Inicial = "A" Then
 Exit Sub
End If
If (Grid_Arti.COL = 5 Or Grid_Arti.COL = 6) Or Grid_Arti.COL >= 11 And Grid_Arti.COL <= 17 Then
 loc_unid = "A"
 ElGrid_KeyDown Grid_Arti, TEXTOVAR, KeyCode
 loc_unid = ""
End If
End Sub

Private Sub Grid_Arti_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  If Txt_key.Visible = True Then
     Txt_key.SetFocus
  End If
  If PROV.Visible = True Then
     PROV.SetFocus
  End If
  If lineas.Visible = True Then
     lineas.SetFocus
  End If
  Exit Sub
End If
If KeyAscii = 13 Then
   'grid_det.Col = 4
   grid_det.Enabled = True
   cmdcompras.Enabled = True
   LLENA_DETALLE
   If Flag_Inicial = "X" Then
     grid_det.Enabled = False
     cmdcompras.Enabled = False
     Flag_Inicial = "A"
     Grid_Arti.SetFocus
     Exit Sub
   End If
   loc_fila = Grid_Arti.Row
   loc_colum = Grid_Arti.COL
   For Fila = 1 To 6
     Grid_Arti.COL = Fila
     Grid_Arti.CellForeColor = QBColor(15)
     Grid_Arti.CellBackColor = QBColor(1)
   Next Fila
   'Grid_Arti.Enabled = False
   'grid_det.Clear
   
   Flag_Inicial = "A"
   grid_det.COL = 3
   grid_det.Row = 1
   Flag_Inicial = ""
   grid_compras.Enabled = True
   cmdGrabar.Enabled = True
   grid_det.SetFocus
   Exit Sub
End If
If Flag_Inicial = "A" Then
 Exit Sub
End If
loc_unid = "A"
'ElGrid_KeyPress Grid_Arti, textovar, KeyAscii
loc_unid = ""
End Sub

Private Sub Grid_Arti_LeaveCell()
If Flag_Inicial = "A" Then
 Exit Sub
End If
If Grid_Arti.COL >= 11 Then
 ElGrid_LeaveCell Grid_Arti, TEXTOVAR
End If
If LK_EMP = "PIU" And (Grid_Arti.COL = 5 Or Grid_Arti.COL = 6) Then
 ElGrid_LeaveCell Grid_Arti, TEXTOVAR
End If

End Sub

Private Sub grid_det_Click()
If loc_flag_precios <> "A" Then
 Exit Sub
End If
ElGrid_Click grid_det, TEXTOVAR
End Sub

Private Sub grid_det_DblClick()
If loc_flag_precios <> "A" Then
 Exit Sub
End If
If Flag_Inicial = "A" Then
 Exit Sub
End If
If grid_det.COL = 3 Then Exit Sub
ElGrid_DblClick grid_det, TEXTOVAR
End Sub

Private Sub grid_det_EnterCell()
If loc_flag_precios <> "A" Then
 Exit Sub
End If
If Flag_Inicial = "A" Then
 Exit Sub
End If
ElGrid_EnterCell grid_det, TEXTOVAR
End Sub

Private Sub grid_det_GotFocus()
If loc_flag_precios <> "A" Then
 Exit Sub
End If
If Flag_Inicial = "A" Then
 Exit Sub
End If
ElGrid_GotFocus grid_det, TEXTOVAR
'grid_det.Row = loc_fila
'grid_det.Col = loc_colum

End Sub

Private Sub grid_det_KeyDown(KeyCode As Integer, Shift As Integer)
If loc_flag_precios <> "A" Then
 Exit Sub
End If
If grid_det.COL = 3 Then Exit Sub
ElGrid_KeyDown grid_det, TEXTOVAR, KeyCode
End Sub

Private Sub grid_det_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  Grid_Arti.Enabled = True
  LBLARTI.Caption = ""
  grid_compras.Clear
  grid_compras.Enabled = False
  cmdcompras.Enabled = False
  cmdGrabar.Enabled = False
  grid_det.Clear
  grid_det.Cols = 1
  grid_det.Rows = 1
  grid_det.Enabled = False
  For Fila = -6 To -1
     Grid_Arti.COL = Fila * -1
     Grid_Arti.CellForeColor = QBColor(0)
     Grid_Arti.CellBackColor = QBColor(15)
  Next Fila
  fra2.Enabled = True
  grid_det.Enabled = False
  Grid_Arti.SetFocus
  Exit Sub
End If
If loc_flag_precios <> "A" Then
 Exit Sub
End If
If grid_det.COL = 3 Then Exit Sub

ElGrid_KeyPress grid_det, TEXTOVAR, KeyAscii, 13
End Sub

Private Sub grid_det_LeaveCell()
If loc_flag_precios <> "A" Then
 Exit Sub
End If
If Flag_Inicial = "A" Then
 Exit Sub
End If
If Flag_Change <> "A" Then
  If grid_det.COL = 3 Then ' costo base
    If LK_FLAG_CALCULO = "A" Then
      CALCULAR val(grid_det.TextMatrix(grid_det.Row, 3))
    Else
      CALCULAR_OTRO val(grid_det.TextMatrix(grid_det.Row, 3))
    End If
  End If
  If grid_det.COL = 5 Or grid_det.COL = 7 Or grid_det.COL = 9 Or grid_det.COL = 11 Or grid_det.COL = 13 Then ' costo PORCENTAJE
    CALCULAR_POR val(grid_det.TextMatrix(grid_det.Row, grid_det.COL)), grid_det.COL
    unitPrice = val(grid_det.TextMatrix(grid_det.Row, grid_det.COL)) / val(grid_det.TextMatrix(grid_det.Row, 1))
    'If (grid_det.Row = 1 And grid_det.Rows > 2) Then
     
     For i = 1 To grid_det.Rows - 1
      grid_det.TextMatrix(i, grid_det.COL) = Format(unitPrice * val(grid_det.TextMatrix(i, 1)), "0.0000")
      'grid_det.Row = i
      CALCULAR_PORCOL val(grid_det.TextMatrix(i, grid_det.COL)), grid_det.COL, i
     Next i
    'End If
  End If
  If grid_det.COL = 4 Or grid_det.COL = 6 Or grid_det.COL = 8 Or grid_det.COL = 10 Or grid_det.COL = 12 Then ' costo PORCENTAJE
    CALCULAR_PRE val(grid_det.TextMatrix(grid_det.Row, grid_det.COL)), grid_det.COL
  End If
  Flag_Change = "A"
End If

ElGrid_LeaveCell grid_det, TEXTOVAR
End Sub

Private Sub LINEAS_Click()
Dim wpos As Integer
Dim WFAMI2 As Integer
If Trim(lineas.Text) = "" Then
 listapro.Clear
 Exit Sub
End If
Screen.MousePointer = 11
wpos = listapro.ListIndex
WFAMI2 = val(Trim(Right(lineas.Text, 6)))
LLENADO_SUBFAM WFAMI2
On Error GoTo sigue
listapro.ListIndex = wpos
Screen.MousePointer = 0
Exit Sub
sigue:
Resume Next

End Sub

Private Sub lineas_DblClick()
lineas_KeyPress 13
End Sub

Private Sub lineas_GotFocus()
If loc_unid = "A" Then
 Exit Sub
End If
If Grid_Arti.Enabled = False Then
 Exit Sub
End If
LBLARTI.Caption = ""
grid_det.Cols = 1
grid_det.Rows = 1
cmdGrabar.Enabled = False
grid_det.Clear
grid_det.Enabled = False
Grid_Arti.Clear
Grid_Arti.Rows = 1
Grid_Arti.Cols = 1
Grid_Arti.Enabled = False
grid_compras.Rows = 1
grid_compras.Cols = 1
grid_compras.Clear
grid_compras.Enabled = False
cmdcompras.Enabled = False

End Sub

Private Sub lineas_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  Exit Sub
End If
CmdProcesa_Click
End Sub


Private Sub listapro_DblClick()
listapro_KeyPress 13
End Sub

Private Sub listapro_GotFocus()
If loc_unid = "A" Then
 Exit Sub
End If
If Grid_Arti.Enabled = False Then
 Exit Sub
End If
LBLARTI.Caption = ""
grid_det.Cols = 1
grid_det.Rows = 1
cmdGrabar.Enabled = False
grid_det.Clear
grid_det.Enabled = False
Grid_Arti.Clear
Grid_Arti.Rows = 1
Grid_Arti.Cols = 1
Grid_Arti.Enabled = False
grid_compras.Rows = 3
grid_compras.Enabled = False
cmdcompras.Enabled = False

End Sub

Private Sub listapro_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Exit Sub
End If
CmdProcesa_Click
End Sub

Private Sub ListUnidad_DblClick()
ListUnidad_KeyPress 13
End Sub

Private Sub ListUnidad_KeyPress(KeyAscii As Integer)
On Error GoTo SALE
If KeyAscii = 27 Then
 Grid_Arti.SetFocus
 ListUnidad.Visible = False
 loc_unid = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
If Trim(ListUnidad.Text) = "" Then
 Exit Sub
End If
 Dim wsUnidad As String * 15
 Dim wsEquiv As Currency
 Dim wsCosto As Currency
 loc_unid = "A"
 preuni_llave2.MoveFirst
 If preuni_llave2.EOF Then
  Exit Sub
 End If
 Screen.MousePointer = 11
 Do Until preuni_llave2.EOF
  If preuni_llave2!PRE_FLAG_UNIDAD = "A" Then
    preuni_llave2.Edit
    preuni_llave2!PRE_FLAG_UNIDAD = " "
    preuni_llave2.Update
  End If
  If preuni_llave2!pre_secuencia = val(Right(ListUnidad.Text, 4)) Then
    preuni_llave2.Edit
    preuni_llave2!PRE_FLAG_UNIDAD = "A"
    preuni_llave2.Update
    Grid_Arti.TextMatrix(Grid_Arti.Row, 1) = Trim(preuni_llave2!pre_unidad)
    Grid_Arti.TextMatrix(Grid_Arti.Row, 2) = Format(val(Grid_Arti.TextMatrix(Grid_Arti.Row, 9)) * preuni_llave2!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Grid_Arti.Row, 3) = Format(val(Grid_Arti.TextMatrix(Grid_Arti.Row, 10)) * preuni_llave2!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Grid_Arti.Row, 4) = Format(preuni_llave2!PRE_COSTO, "0.0000")
    Grid_Arti.TextMatrix(Grid_Arti.Row, 5) = preuni_llave2!PRE_POR1
    Grid_Arti.TextMatrix(Grid_Arti.Row, 6) = preuni_llave2!PRE_PRE1
    Grid_Arti.TextMatrix(Grid_Arti.Row, 7) = preuni_llave2!PRE_codart
    Grid_Arti.TextMatrix(Grid_Arti.Row, 8) = preuni_llave2!PRE_EQUIV
    Grid_Arti.TextMatrix(Grid_Arti.Row, 11) = Format((Grid_Arti.TextMatrix(Grid_Arti.Row, 19) / preuni_llave2!PRE_EQUIV), "0.0000")
    Grid_Arti.TextMatrix(Grid_Arti.Row, 12) = Format((Grid_Arti.TextMatrix(Grid_Arti.Row, 20) / preuni_llave2!PRE_EQUIV), "0.0000")
    
  End If
 preuni_llave2.MoveNext
 Loop
 Grid_Arti.SetFocus
 Screen.MousePointer = 0
 ListUnidad.Visible = False
loc_unid = ""
Exit Sub
SALE:
 MsgBox Err.Description + ", Intente Nuevamente ", 48, Pub_Titulo
End Sub

Private Sub ListUnidad_LostFocus()
ListUnidad.Visible = False
End Sub

Private Sub ListView1_DblClick()
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.item(loc_key).Text) & " "
 txt_key_KeyPress 13

End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.item(loc_key).Selected = True
 ListView1.ListItems.item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 Txt_key.Text = Trim(ListView1.ListItems.item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
txt_key_KeyPress 13
End Sub


Private Sub Option1_Click(Index As Integer)
If option1(0).Value = True Then
   lblopcion.Caption = ""
   Txt_key.Text = ""
   Txt_key.Visible = False
   lineas.Visible = False
   listapro.Visible = False
   PROV.Visible = False
   Label1.Visible = False
   Label1.Caption = "Codigo :"
   Label2.Visible = False
   CmdProcesa.Visible = False
   txtdato.Visible = True
   txtdato.SetFocus
ElseIf option1(1).Value = True Then
   lblopcion.Caption = ""
   Txt_key.Text = ""
   lineas.Visible = True
   listapro.Visible = True
   CmdProcesa.Visible = True
   txtdato.Visible = False
   If listapro.ListCount > 0 Then
     listapro.ListIndex = -1
   End If
   If lineas.ListCount > 0 Then
      lineas.ListIndex = -1
   End If
   PROV.Visible = False
   Label1.Visible = True
   Label1.Caption = "Familia :"
   Label2.Visible = True
   lineas.SetFocus
ElseIf option1(2).Value = True Then
   For i = 0 To PROV.ListCount - 1
      PROV.Selected(i) = False
   Next i
   lblopcion.Caption = ""
   Txt_key.Text = ""
   lineas.Visible = False
   listapro.Visible = False
   PROV.Visible = True
   Label1.Caption = "Proveedor"
   Label1.Visible = True
   Label2.Visible = False
   Txt_key.Visible = False
   PROV.SetFocus
End If

End Sub

Private Sub Option1_GotFocus(Index As Integer)
If Grid_Arti.Enabled = False Then
 Exit Sub
End If
LBLARTI.Caption = ""
grid_det.Cols = 1
grid_det.Rows = 1
cmdGrabar.Enabled = False
grid_det.Clear
grid_det.Enabled = False
Grid_Arti.Clear
Grid_Arti.Rows = 1
Grid_Arti.Cols = 1
Grid_Arti.Enabled = False
grid_compras.Rows = 1
grid_compras.Cols = 1
grid_compras.Clear
grid_compras.Enabled = False
cmdcompras.Enabled = False
End Sub

Private Sub preliminar_Click()
Dim cam_llave As rdoResultset
Dim CAM_PRECIO As rdoQuery
Dim WS_FAMI As String
Dim WVER1 As String
Dim WVER2 As String
gridp.Cols = 3
gridp.Rows = 1
gridp.ColWidth(0) = 1500
gridp.ColWidth(1) = 900
gridp.ColWidth(2) = 900
gridp.TextMatrix(0, 0) = "Producto"
gridp.TextMatrix(0, 1) = "Mayor."
gridp.TextMatrix(0, 2) = "Minor."
gridp.ColWidth(1) = 900
gridp.ColWidth(2) = 900

WS_FAMI = Trim(Right(fami.Text, 8))

WVER1 = ""
For Fila = 0 To subfami.ListCount - 1
  subfami.ListIndex = Fila
  If subfami.Selected(Fila) Then
    WVER1 = WVER1 & " ART_SUBFAM = " & Trim(Right(subfami.Text, 8)) & " OR"
  End If
Next Fila
If WVER1 <> "" Then
    WVER1 = Mid(WVER1, 1, Len(WVER1) - 3)
End If

WVER2 = ""
For Fila = 0 To grupo.ListCount - 1
  grupo.ListIndex = Fila
  If grupo.Selected(Fila) And Trim(grupo.Text) <> "" Then
    WVER2 = WVER2 & " ART_SUBGRU = " & Trim(Right(grupo.Text, 8)) & " OR"
  End If
Next Fila
If WVER2 <> "" Then
    WVER2 = Mid(WVER2, 1, Len(WVER2) - 3)
End If


pub_cadena = "SELECT * FROM ARTI,PRECIOS WHERE (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND ART_CODCIA = ? AND ART_FAMILIA = '" & WS_FAMI & "' "
If WVER1 <> "" Then
 pub_cadena = pub_cadena & " AND (" & WVER1 & ") "
End If
If WVER2 <> "" Then
 pub_cadena = pub_cadena & " AND (" & WVER2 & ") "
End If

'Debug.Print pub_cadena
Set CAM_PRECIO = CN.CreateQuery("", pub_cadena)
CAM_PRECIO(0) = 0
Set cam_llave = CAM_PRECIO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
CAM_PRECIO(0) = LK_CODCIA
cam_llave.Requery
Do Until cam_llave.EOF
  gridp.Rows = gridp.Rows + 1
  gridp.RowHeight(gridp.Rows - 1) = 285
  gridp.TextMatrix(gridp.Rows - 1, 0) = cam_llave!art_nombre
  gridp.TextMatrix(gridp.Rows - 1, 1) = cam_llave!PRE_PRE1
  gridp.TextMatrix(gridp.Rows - 1, 2) = cam_llave!PRE_PRE2
  cam_llave.MoveNext
Loop
gridp.Visible = True
End Sub

Private Sub propre_Click()
Dim cam_llave As rdoResultset
Dim CAM_PRECIO As rdoQuery
Dim WS_FAMI As String
Dim WVER1 As String
Dim WVER2 As String
pub_mensaje = "Procesar Cambio de Lista de Precios... ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

WS_FAMI = Trim(Right(fami.Text, 8))

WVER1 = ""
For Fila = 0 To subfami.ListCount - 1
  subfami.ListIndex = Fila
  If subfami.Selected(Fila) Then
    WVER1 = WVER1 & " ART_SUBFAM = " & Trim(Right(subfami.Text, 8)) & " OR"
  End If
Next Fila
If WVER1 <> "" Then
    WVER1 = Mid(WVER1, 1, Len(WVER1) - 3)
End If

WVER2 = ""
For Fila = 0 To grupo.ListCount - 1
  grupo.ListIndex = Fila
  If grupo.Selected(Fila) Then
    WVER2 = WVER2 & " ART_SUBGRU = " & Trim(Right(grupo.Text, 8)) & " OR"
  End If
Next Fila
If WVER2 <> "" Then
    WVER2 = Mid(WVER2, 1, Len(WVER2) - 3)
End If


pub_cadena = "SELECT * FROM ARTI,PRECIOS WHERE (ART_CODCIA = PRE_CODCIA) AND (ART_KEY = PRE_CODART) AND ART_CODCIA = ? AND ART_FAMILIA = '" & WS_FAMI & "' "
If WVER1 <> "" Then
 pub_cadena = pub_cadena & " AND (" & WVER1 & ") "
End If
If WVER2 <> "" Then
 pub_cadena = pub_cadena & " AND (" & WVER2 & ") "
End If

'Debug.Print pub_cadena
Set CAM_PRECIO = CN.CreateQuery("", pub_cadena)
CAM_PRECIO(0) = 0
Set cam_llave = CAM_PRECIO.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
CAM_PRECIO(0) = LK_CODCIA
cam_llave.Requery
Do Until cam_llave.EOF
  SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODART = cam_llave!ART_KEY
   PUB_SECUEN = 0
   LEER_PRE_LLAVE
   pre_llave.Edit
   If val(pre1.Text) <> 0 Then pre_llave!PRE_PRE1 = val(pre1.Text)
   If val(pre2.Text) <> 0 Then pre_llave!PRE_PRE2 = val(pre2.Text)
   If val(pre3.Text) <> 0 Then pre_llave!PRE_PRE3 = val(pre3.Text)
   If val(pre4.Text) <> 0 Then pre_llave!PRE_PRE4 = val(pre4.Text)
   If val(pre5.Text) <> 0 Then pre_llave!PRE_PRE5 = val(pre5.Text)
   pre_llave.Update
 cam_llave.MoveNext
Loop
MsgBox "Proceso Terminado ", 48, Pub_Titulo
End Sub

Private Sub PROV_GotFocus()
If loc_unid = "A" Then
 Exit Sub
End If
If Grid_Arti.Enabled = False Then
 Exit Sub
End If
LBLARTI.Caption = ""
grid_det.Cols = 1
grid_det.Rows = 1
cmdGrabar.Enabled = False
grid_det.Clear
grid_det.Enabled = False
Grid_Arti.Clear
Grid_Arti.Rows = 1
Grid_Arti.Cols = 1
Grid_Arti.Enabled = False
grid_compras.Rows = 1
grid_compras.Cols = 1
grid_compras.Clear
grid_compras.Enabled = False
cmdcompras.Enabled = False
End Sub

Private Sub PROV_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
Exit Sub
End If
CmdProcesa_Click
End Sub

Private Sub salir_Click()
 Unload FrmPreUnidad
End Sub

Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)
If grid_det.Enabled Then
 If KeyCode = 40 Or KeyCode = 37 Or KeyCode = 39 Or KeyCode = 38 Then
  If grid_det.COL = 4 Or grid_det.COL = 6 Or grid_det.COL = 8 Or grid_det.COL = 10 Or grid_det.COL = 12 Then
    If val(TEXTOVAR.Text) > 999.99 Or val(TEXTOVAR.Text) < 0 Then
       Azul TEXTOVAR, TEXTOVAR
       Beep
       KeyAscii = 0
      Exit Sub
     End If
   End If
 End If
 TEXTO_KeyDown grid_det, TEXTOVAR, KeyCode, 13
Else
 loc_unid = "A"
 TEXTO_KeyDown Grid_Arti, TEXTOVAR, KeyCode
End If
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If grid_det.COL = 4 Or grid_det.COL = 6 Or grid_det.COL = 8 Or grid_det.COL = 10 Or grid_det.COL = 12 Then
  If val(TEXTOVAR.Text) > 999.99 Or val(TEXTOVAR.Text) < 0 Then
   Azul TEXTOVAR, TEXTOVAR
   Beep
   KeyAscii = 0
   Exit Sub
  End If
End If
End If
If grid_det.Enabled Then
 TEXTO_KeyPress grid_det, TEXTOVAR, KeyAscii, 13, 3, 2, 4, 2, 5, 2, 6, 2, 7, 2, 8, 2, 9, 2, 10, 2, 11, 2, 12, 2, 13, 2
Else
 loc_unid = "A"
 TEXTO_KeyPress Grid_Arti, TEXTOVAR, KeyAscii, , 11, 2, 12, 2, 13, 2, 14, 2, 15, 2, 16, 2, 17, 2, 18, 2
End If
End Sub

Private Sub textovar_LostFocus()
If grid_det.Enabled Then
 TEXTO_LosFocus grid_det, TEXTOVAR
Else
 loc_unid = "A"
 TEXTO_LosFocus Grid_Arti, TEXTOVAR
End If
End Sub

Private Sub Txt_key_Change()
If Txt_key.Text = "" Then
 lblopcion.Caption = ""
 VAR_ACTIVAR = 0
End If
End Sub

Private Sub txt_key_GotFocus()
PRO_GOTFOCUS
End Sub

Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And Txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.item(loc_key).Selected = True
  ListView1.ListItems.item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
  ListView1.ListItems.item(loc_key).Selected = True
  ListView1.ListItems.item(loc_key).EnsureVisible
  Txt_key.Text = Trim(ListView1.ListItems.item(loc_key).Text)
  DoEvents
  Txt_key.SelStart = Len(Txt_key.Text)
  DoEvents
fin:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 Txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 PUB_KEY = val(Txt_key.Text)
 If Len(Txt_key.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(Txt_key.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = Txt_key.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul Txt_key, Txt_key
       GoTo fin
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    lblopcion.Caption = Trim(art_LLAVE!art_nombre)
    CmdProcesa_Click
    ListView1.Visible = False
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And Txt_key.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = Txt_key.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul Txt_key, Txt_key
       Exit Sub
     End If
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     lblopcion.Caption = Trim(art_llave_alt!art_nombre)
     CmdProcesa_Click
     ListView1.Visible = False
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView1.ListItems.item(loc_key).Text)
    If Trim(UCase(Txt_key.Text)) = Left(valor, Len(Trim(Txt_key.Text))) And Len(Trim(Txt_key.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        Txt_key.Text = Trim(ListView1.ListItems.item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       Txt_key.Text = Trim(ListView1.ListItems.item(loc_key).SubItems(1))
      Else
       Txt_key.Text = Trim(ListView1.ListItems.item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = val(ListView1.ListItems.item(loc_key).SubItems(1))
      Else
       PUB_KEY = Txt_key.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul Txt_key, Txt_key
        Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      lblopcion.Caption = Trim(art_LLAVE!art_nombre)
      CmdProcesa_Click
      ListView1.Visible = False
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView1.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul Txt_key, Txt_key
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim var
Dim ws_codcia As String * 2
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(Txt_key.Text) = 0 Or Txt_key.Text = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If Txt_key.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf Txt_key.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(Txt_key.Text, 1) = "*" Then
   Txt_key.Text = Mid(Txt_key.Text, 2, Len(Txt_key.Text))
   Txt_key.SelStart = Len(Txt_key.Text)
  End If
Else
 If Len(Txt_key.Text) = 0 Or IsNumeric(Txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(Txt_key.Text) = 1 Then
    var = Asc(Txt_key.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    Else
       var = Chr(var)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND  (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CALIDAD = " & wnormal & " AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      ''archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK , PRE_EQUIV FROM ARTI, ARTICULO, PRECIOS  WHERE  (ART_KEY = PRE_CODART) AND (ART_CODCIA = PRE_CODCIA) AND (PRE_FLAG_UNIDAD ='A') AND (ART_KEY = ARM_CODART) AND (ART_CODCIA = ARM_CODCIA) AND ART_CALIDAD = " & wnormal & " AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & Txt_key.Text & "' AND  '" & var & "' ORDER BY ART_NOMBRE"
      
    End If
     PROC_LISVIEW ListView1
     loc_key = 0
     If ListView1.Visible Then
      loc_key = 1
     End If
     Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(Txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

End Sub



Public Sub CABEZA_ARTI()
Dim parte1
Dim parte2
Grid_Arti.FixedRows = 2
Grid_Arti.Cols = 21
Grid_Arti.ColWidth(0) = 6500 ' nombre
Grid_Arti.ColWidth(1) = 700 ' unidad
Grid_Arti.ColWidth(2) = 1 '700  'c.prom.
Grid_Arti.ColWidth(3) = 600 '700  ' c. Repos.
Grid_Arti.ColWidth(4) = 1  ' cos. base
Grid_Arti.ColWidth(5) = 1000   'Precio 1
Grid_Arti.ColWidth(6) = 1000    'Precio 2
Grid_Arti.ColWidth(7) = 1    'codigo articulo
Grid_Arti.ColWidth(8) = 1    'EQUIVALENCIA
Grid_Arti.ColWidth(9) = 1    'costo promedio
Grid_Arti.ColWidth(10) = 1    'costo reposicion
Grid_Arti.Visible = True
 Grid_Arti.ColWidth(11) = 1
 Grid_Arti.ColWidth(12) = 1
 Grid_Arti.ColWidth(13) = 1
 Grid_Arti.ColWidth(14) = 1
 Grid_Arti.ColWidth(15) = 1
 Grid_Arti.ColWidth(16) = 1
 Grid_Arti.ColWidth(17) = 1
 Grid_Arti.ColWidth(18) = 1
 Grid_Arti.ColWidth(19) = 1
 Grid_Arti.ColWidth(20) = 1
 
If loc_flag_stock = "A" Then
 Grid_Arti.ColWidth(11) = 1  'STOCK MINIMO
 Grid_Arti.ColWidth(12) = 1  'STOCK MAXIMO
End If
If loc_flag_por = "A" Then
 Grid_Arti.ColWidth(13) = 1  '% Comis. 1
 Grid_Arti.ColWidth(14) = 1  '% Comis. 2
 Grid_Arti.ColWidth(15) = 1  '% Comis. 3
 Grid_Arti.ColWidth(16) = 1  '% Comis. 4
 Grid_Arti.ColWidth(17) = 1  '% Comis. 5
End If
If loc_flag_stock = "A" Then
 Grid_Arti.ColWidth(18) = 1  '% STOCK INICIAL
End If

Grid_Arti.RowHeight(0) = 300
Grid_Arti.TextMatrix(0, 0) = ""
Grid_Arti.COL = 0
Grid_Arti.Row = 0
Grid_Arti.CellFontName = "Arial"
Grid_Arti.TextMatrix(0, 0) = "ENTER = Editar Precios"
Grid_Arti.CellTextStyle = 4
Grid_Arti.CellFontSize = 7
Grid_Arti.Row = 1
Grid_Arti.CellTextStyle = 3
Grid_Arti.CellFontBold = True

Grid_Arti.TextMatrix(1, 0) = "ARTICULOS "
Grid_Arti.TextMatrix(0, 1) = "Unidad"
Grid_Arti.TextMatrix(1, 1) = "Activa"
Grid_Arti.TextMatrix(0, 2) = "Costo"
Grid_Arti.TextMatrix(1, 2) = "Prom."
Grid_Arti.TextMatrix(0, 3) = "Costo"
Grid_Arti.TextMatrix(1, 3) = "Repos."
Grid_Arti.TextMatrix(0, 4) = "Costo"
Grid_Arti.TextMatrix(1, 4) = " Base"
Grid_Arti.TextMatrix(0, 11) = "Stock"
Grid_Arti.TextMatrix(1, 11) = "Minimo"
Grid_Arti.TextMatrix(0, 12) = "Stock "
Grid_Arti.TextMatrix(1, 12) = "Maximo"
Grid_Arti.TextMatrix(0, 13) = "(%)Coms."
Grid_Arti.TextMatrix(1, 13) = loc_comis(1)
Grid_Arti.TextMatrix(0, 14) = "(%)Coms."
Grid_Arti.TextMatrix(1, 14) = loc_comis(2)
Grid_Arti.TextMatrix(0, 15) = "(%)Coms."
Grid_Arti.TextMatrix(1, 15) = loc_comis(3)
Grid_Arti.TextMatrix(0, 16) = "(%)Coms."
Grid_Arti.TextMatrix(1, 16) = loc_comis(4)
Grid_Arti.TextMatrix(0, 17) = "(%)Coms."
Grid_Arti.TextMatrix(1, 17) = loc_comis(5)
Grid_Arti.TextMatrix(0, 18) = "Stock"
Grid_Arti.TextMatrix(1, 18) = "Inicial"
parte1 = ""
parte2 = ""
On Error GoTo sigue
If LK_EMP <> "PIU" Then
 parte1 = Left(Label3(0).Caption, Int(Len(Label3(0).Caption) / 2))
 parte2 = Mid(Label3(0).Caption, Int(Len(Label3(0).Caption) / 2) + 1, Len(Label3(0).Caption))
On Error GoTo 0
Grid_Arti.TextMatrix(0, 5) = "PRECIO"
Grid_Arti.COL = 5
Grid_Arti.Row = 0
Grid_Arti.CellTextStyle = 3
Grid_Arti.CellAlignment = 1
Grid_Arti.CellFontBold = True
Grid_Arti.TextMatrix(1, 5) = "MENOR"
Grid_Arti.TextMatrix(0, 6) = "PRECIO"
Grid_Arti.COL = 6
Grid_Arti.CellFontBold = True
Grid_Arti.CellTextStyle = 3
Grid_Arti.CellAlignment = 1
Grid_Arti.TextMatrix(1, 6) = "MAYOR"
Else
Grid_Arti.TextMatrix(1, 5) = Label3(0).Caption
Grid_Arti.TextMatrix(1, 6) = Label3(1).Caption

End If
Exit Sub
sigue:
Resume Next
End Sub
Public Sub CABEZA_DET()
Flag_Inicial = "A"
grid_det.FixedRows = 1
grid_det.Cols = 14
grid_det.Cols = 15
grid_det.ColWidth(0) = 1100 ' unidad
grid_det.ColWidth(1) = 600 ' equivalencia
grid_det.ColWidth(2) = 700  'c. Repos.
grid_det.ColWidth(3) = 700  ' cos. base
grid_det.ColWidth(4) = 550   ' % 1
grid_det.ColWidth(5) = 850   ' p 1
grid_det.ColWidth(6) = 550   ' % 2
grid_det.ColWidth(7) = 850   ' p 2
grid_det.ColWidth(8) = 550   ' % 3
grid_det.ColWidth(9) = 850   ' p 3
grid_det.ColWidth(10) = 550   ' % 4
grid_det.ColWidth(11) = 850   ' p 4
grid_det.ColWidth(12) = 550   ' % 5
grid_det.ColWidth(13) = 850   ' p 5
grid_det.ColWidth(14) = 1   ' x costo reposicion
grid_det.TextMatrix(0, 0) = "Unidades"
grid_det.TextMatrix(0, 1) = "Equiv."
grid_det.TextMatrix(0, 2) = "C.Rep."
grid_det.TextMatrix(0, 3) = " PROM."
grid_det.TextMatrix(0, 4) = "( % )  "
grid_det.TextMatrix(0, 5) = "  Valor"
grid_det.TextMatrix(0, 6) = "( % )  "
grid_det.TextMatrix(0, 7) = "  Valor"
grid_det.TextMatrix(0, 8) = "( % )  "
grid_det.TextMatrix(0, 9) = "  Valor"
grid_det.TextMatrix(0, 10) = "( % )  "
grid_det.TextMatrix(0, 11) = "  Valor"
grid_det.TextMatrix(0, 12) = "( % )  "
grid_det.TextMatrix(0, 13) = "  Valor"
grid_det.TextMatrix(0, 14) = "Cos.Repos."

Flag_Inicial = ""
End Sub

Public Sub LLENA_DETALLE()
Dim WSPOR As Currency
PRE_UNIDADES2(0) = LK_CODCIA
PRE_UNIDADES2(1) = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 7))
preuni_llave2.Requery
If preuni_llave2.EOF Then
 MsgBox "Error de Unidades NO existe......", 48, Pub_Titulo
 Flag_Inicial = "X"
 Exit Sub
End If
Fila = 0
grid_det.Clear
grid_det.Rows = 3
grid_det.Cols = 14
grid_det.FixedRows = 1
grid_det.FixedCols = 2
CABEZA_DET
grid_det.ColWidth(3) = 1  ' cos. base
grid_det.ColWidth(0) = 1800  ' cos. base
grid_det.ColAlignment(0) = 1
LBLARTI.Caption = Trim(Grid_Arti.TextMatrix(Grid_Arti.Row, 0))
Flag_Inicial = "A"
Do Until preuni_llave2.EOF
   Fila = Fila + 1
   grid_det.Rows = Fila + 1
   grid_det.RowHeight(Fila) = 285
   grid_det.Row = Fila
   'grid_det.Col = 0
   'grid_det.CellFontBold = True
   grid_det.TextMatrix(Fila, 0) = Trim(preuni_llave2!pre_unidad)
   grid_det.TextMatrix(Fila, 1) = preuni_llave2!PRE_EQUIV
   'grid_det.TextMatrix(fila, 2) = Format(Val(Grid_Arti.TextMatrix(Grid_Arti.Row, 10)) * preuni_llave2!PRE_EQUIV, "0.00")
   grid_det.TextMatrix(Fila, 2) = Format((Nulo_Valor0(preuni_llave2!PRE_COSTO) * ((100 + LK_IGV) / 100)), "0.0000")
   grid_det.TextMatrix(Fila, 3) = Format(val(Grid_Arti.TextMatrix(Grid_Arti.Row, 9)) * preuni_llave2!PRE_EQUIV, "0.0000")
   grid_det.COL = 4
   grid_det.CellForeColor = QBColor(9)
   If val(grid_det.TextMatrix(Fila, 2)) <> 0 Then WSPOR = (Nulo_Valor0(preuni_llave2!PRE_PRE1) * 100) / val(grid_det.TextMatrix(Fila, 2)) - 100
   grid_det.TextMatrix(Fila, 4) = Format(WSPOR, "0.0000") ' Nulo_Valor0(preuni_llave2!PRE_POR1)
   grid_det.TextMatrix(Fila, 5) = Nulo_Valor0(preuni_llave2!PRE_PRE1)
   grid_det.COL = 6
   grid_det.CellForeColor = QBColor(9)
   If val(grid_det.TextMatrix(Fila, 2)) <> 0 Then WSPOR = (Nulo_Valor0(preuni_llave2!PRE_PRE2) * 100) / val(grid_det.TextMatrix(Fila, 2)) - 100
   grid_det.TextMatrix(Fila, 6) = Format(WSPOR, "0.0000")
   grid_det.TextMatrix(Fila, 7) = Nulo_Valor0(preuni_llave2!PRE_PRE2)
   grid_det.COL = 8
   grid_det.CellForeColor = QBColor(9)
   If val(grid_det.TextMatrix(Fila, 2)) <> 0 Then WSPOR = (Nulo_Valor0(preuni_llave2!PRE_PRE3) * 100) / val(grid_det.TextMatrix(Fila, 2)) - 100
   grid_det.TextMatrix(Fila, 8) = Format(WSPOR, "0.0000")
   grid_det.TextMatrix(Fila, 9) = Nulo_Valor0(preuni_llave2!PRE_PRE3)
   grid_det.COL = 10
   grid_det.CellForeColor = QBColor(9)
   If val(grid_det.TextMatrix(Fila, 2)) <> 0 Then WSPOR = (Nulo_Valor0(preuni_llave2!PRE_PRE4) * 100) / val(grid_det.TextMatrix(Fila, 2)) - 100
   grid_det.TextMatrix(Fila, 10) = Format(WSPOR, "0.0000")
   grid_det.TextMatrix(Fila, 11) = Nulo_Valor0(preuni_llave2!PRE_PRE4)
   grid_det.COL = 12
   grid_det.CellForeColor = QBColor(9)
   If val(grid_det.TextMatrix(Fila, 2)) <> 0 Then WSPOR = (Nulo_Valor0(preuni_llave2!PRE_PRE5) * 100) / val(grid_det.TextMatrix(Fila, 2)) - 100
   grid_det.TextMatrix(Fila, 12) = Format(WSPOR, "0.0000")
   grid_det.TextMatrix(Fila, 13) = Nulo_Valor0(preuni_llave2!PRE_PRE5)
   grid_det.TextMatrix(Fila, 14) = Format(Nulo_Valor0(preuni_llave2!PRE_COSTO_REPO), "0.0000")
   preuni_llave2.MoveNext
Loop
Flag_Inicial = ""
grid_det.Row = 1
grid_det.COL = 3
grid_det.SetFocus
End Sub

Public Sub ElGrid_Click(wsGrid As MSFlexGrid, wsTexto As TextBox)
If wsGrid.CellWidth < 0 Then
Exit Sub
End If
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Visible = False

End Sub
Public Sub ElGrid_EnterCell(wsGrid As MSFlexGrid, wsTexto As TextBox, Optional Bloq1, Optional Bloq2, Optional Bloq3, Optional Bloq4, Optional Bloq5)
If wsGrid.CellWidth < 0 Then
'  wsGrid.SetFocus
 Exit Sub
End If
wsGrid.CellBackColor = QBColor(7)
'wsGrid.CellBackColor = QBColor(1)
'wsGrid.CellForeColor = QBColor(15)
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
Flag_Bloq = ""
If Not IsMissing(Bloq1) Then
  If wsGrid.COL = Bloq1 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq2) Then
  If wsGrid.COL = Bloq2 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq3) Then
  If wsGrid.COL = Bloq3 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq4) Then
  If wsGrid.COL = Bloq4 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq5) Then
  If wsGrid.COL = Bloq5 Then
    Flag_Bloq = "A"
  End If
End If

End Sub
Public Sub ElGrid_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode)
Flag_F2 = ""
If Flag_Bloq = "A" Then
  wsKeyCode = 0
  Exit Sub
End If

If wsKeyCode <> 113 Then
 Exit Sub
End If
If wsTexto.Visible = False Then
  Flag_F2 = "A"
  ElGrid_DblClick wsGrid, wsTexto
End If
End Sub
Public Sub ElGrid_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii, Optional SaltaCol)
If wsKeyAscii = 27 Then
 Exit Sub
End If
If wsKeyAscii = 9 Or wsKeyAscii = 13 Then
  If Not IsMissing(SaltaCol) Then
    If wsGrid.COL = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       wsGrid.COL = wsGrid.FixedCols
       Exit Sub
    End If
    
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
    wsGrid.COL = wsGrid.COL + 1
  End If
  Exit Sub
End If
If Flag_Bloq = "A" Then
 wsKeyAscii = 0
 Exit Sub
End If

Dim cade


wsTexto.Text = ""
wsTexto.Visible = True
cade = UCase(Chr(wsKeyAscii))
'wsTexto.text = cade
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
Flag_Change = "A"
'cade = Chr(wsKeyAscii)
SendKeys cade, True
wsTexto.SelStart = Len(wsTexto)

'wsTexto.text = ""
'wsTexto.Visible = True
'cade = Chr(wsKeyAscii)
'wsTexto.text = cade
'If wsTexto.Enabled = True And wsTexto.Visible = True Then
'   wsTexto.SetFocus
'   wsTexto.SelStart = 0
'   wsTexto.SelLength = Len(wsTexto)
'End If
'Flag_Change = "A"
'cade = Chr(wsKeyAscii)
'SendKeys cade, True
'wsTexto.SelStart = Len(wsTexto)







End Sub
Private Sub ElGrid_LeaveCell(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Consis = "A" Then
 'wsTexto.FontBold = True
' wsTexto.ForeColor = QBColor(12)
 
 wsTexto.Visible = True
 If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
 End If
 Exit Sub
End If
wsGrid.CellBackColor = QBColor(15)
If Left(Trim(wsGrid.Text), 1) = "-" Then
' wsGrid.CellForeColor = QBColor(12)
' wsGrid.CellBackColor = QBColor(15)
Else
' wsGrid.CellBackColor = QBColor(15)
' wsGrid.CellForeColor = QBColor(0)
End If
'wsGrid.CellFontBold = False
End Sub
Private Sub ElGrid_DblClick(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Bloq = "A" Then
  Exit Sub
End If
'wsTexto.FontBold = True
'wsTexto.ForeColor = QBColor(12)
wsTexto.Visible = True
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
End Sub
Private Sub ElGrid_GotFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Private Sub TEXTO_LosFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Public Sub TEXTO_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode As Integer, Optional SaltaCol)
If wsKeyCode = 40 Or wsKeyCode = 37 Or wsKeyCode = 39 Or wsKeyCode = 38 Then
 If Flag_F2 = "A" Then
   Exit Sub
 End If
 If Flag_Consis = "A" Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
   Beep
   Exit Sub
 End If
 Flag_Change = ""
 If wsGrid.COL >= 13 And wsGrid.COL <= 17 Then
   If val(wsTexto.Text) > 999.99 Then
    Exit Sub
   End If
 End If
 If wsGrid.COL = 18 Then
   If Len(Trim(wsTexto.Text)) > 17 Then
    Exit Sub
   End If
 End If
 If Trim(wsTexto.Text) = "." Or Trim(wsTexto.Text) = "" Then
  wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = "0.0000"
 Else
  wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.0000")
 End If
 
 If FrmPreUnidad.grid_det.Enabled And wsGrid.COL = 14 And wsGrid.Row = 2 Then
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.0000")
    wsGrid.TextMatrix(wsGrid.Row - 1, wsGrid.COL) = Format(wsTexto.Text / wsGrid.TextMatrix(wsGrid.Row, 1), "0.0000")
 End If
 
 If FrmPreUnidad.grid_det.Enabled = False And FrmPreUnidad.Grid_Arti.Enabled Then
   If wsGrid.COL = 5 Or wsGrid.COL = 6 Then
      GRABA_POR val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 99
      GoTo sa
   End If
   If wsGrid.COL = 11 Then
     GRABA_DATO val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 1
   ElseIf wsGrid.COL = 12 Then
     GRABA_DATO val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 2
   ElseIf wsGrid.COL = 13 Then
     GRABA_POR val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 1
   ElseIf wsGrid.COL = 14 Then
     GRABA_POR val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 2
   ElseIf wsGrid.COL = 15 Then
     GRABA_POR val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 3
   ElseIf wsGrid.COL = 16 Then
     GRABA_POR val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 4
   ElseIf wsGrid.COL = 17 Then
     GRABA_POR val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), 5
   End If
 End If
sa:
 Flag_Bloq = ""
 wsTexto.Visible = False
 If wsKeyCode = 40 Then ' ABAJO
  If wsGrid.Row <> wsGrid.Rows - 1 Then
     wsGrid.Row = wsGrid.Row + 1
  End If
 End If
 If wsKeyCode = 38 Then ' arriba
  If wsGrid.Row <> wsGrid.FixedRows Then
     wsGrid.Row = wsGrid.Row - 1
  End If
 End If
 If wsKeyCode = 37 Then ' isquierda
  If wsGrid.COL <> wsGrid.FixedCols Then
     wsGrid.COL = wsGrid.COL - 1
  End If
 End If
 If wsKeyCode = 39 Then ' derecha
  If Not IsMissing(SaltaCol) Then
     If wsGrid.COL = SaltaCol Then
        If wsGrid.Row <> wsGrid.Rows - 1 Then
          wsGrid.Row = wsGrid.Row + 1
          wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
       ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
         'wsGrid.Row = wsGrid.Row + 1
          wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
        End If
     ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
        If FrmPreUnidad.grid_det.Enabled And wsGrid.COL = 14 And wsGrid.Row = 2 Then
          wsGrid.COL = wsGrid.COL
        Else

          wsGrid.COL = wsGrid.FixedCols
        End If
        GoTo wsfinal
     End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
    If LK_EMP = "PIU" And wsGrid.COL = 6 Then
    Else
    wsGrid.COL = wsGrid.COL + 1
    End If
      
  End If
 End If
wsfinal:
 wsTexto.FontBold = False
' wsTexto.ForeColor = QBColor(0)
 wsTexto.Text = ""
 wsGrid.SetFocus
End If
'Exit Sub

End Sub

Public Sub TEXTO_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional SaltaCol, Optional ConsisCol1, Optional ConsisVal1, Optional ConsisCol2, Optional ConsisVal2, Optional ConsisCol3, Optional ConsisVal3, Optional ConsisCol4, Optional ConsisVal4, Optional ConsisCol5, Optional ConsisVal5, Optional ConsisCol6, Optional ConsisVal6, Optional ConsisCol7, Optional ConsisVal7, Optional ConsisCol8, Optional ConsisVal8, Optional ConsisCol9, Optional ConsisVal9, Optional ConsisCol10, Optional ConsisVal10, Optional ConsisCol11, Optional ConsisVal11, Optional ConsisCol12, Optional ConsisVal12)
If wsKeyAscii = 13 Or wsKeyAscii = 9 Then
  Flag_F2 = ""
  TEXTO_KeyDown wsGrid, wsTexto, 39, SaltaCol
  Exit Sub
End If
If wsKeyAscii = 27 Then
  ElGrid_Click wsGrid, wsTexto
  Flag_Change = "A"
  wsGrid.SetFocus
End If
If Not IsMissing(ConsisCol1) And Not IsMissing(ConsisVal1) Then
  If wsGrid.COL = ConsisCol1 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal1, ConsisCol1
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol2) And Not IsMissing(ConsisVal2) Then
  If wsGrid.COL = ConsisCol2 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal2, ConsisCol2
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol3) And Not IsMissing(ConsisVal3) Then
  If wsGrid.COL = ConsisCol3 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal3, ConsisCol3
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol4) And Not IsMissing(ConsisVal4) Then
  If wsGrid.COL = ConsisCol4 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal4, ConsisCol4
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol5) And Not IsMissing(ConsisVal5) Then
  If wsGrid.COL = ConsisCol5 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal5, ConsisCol5
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol6) And Not IsMissing(ConsisVal6) Then
  If wsGrid.COL = ConsisCol6 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal6, ConsisCol6
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol7) And Not IsMissing(ConsisVal7) Then
  If wsGrid.COL = ConsisCol7 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal7, ConsisCol7
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol8) And Not IsMissing(ConsisVal8) Then
  If wsGrid.COL = ConsisCol8 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal8, ConsisCol8
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol9) And Not IsMissing(ConsisVal9) Then
  If wsGrid.COL = ConsisCol9 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal9, ConsisCol9
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol10) And Not IsMissing(ConsisVal10) Then
  If wsGrid.COL = ConsisCol10 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal10, ConsisCol10
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol11) And Not IsMissing(ConsisVal11) Then
  If wsGrid.COL = ConsisCol11 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal11, ConsisCol11
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol12) And Not IsMissing(ConsisVal12) Then
  If wsGrid.COL = ConsisCol12 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal12, ConsisCol12
   Exit Sub
  End If
End If

End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static valor
  Dim car As String
  Flag_Consis = ""
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "-" Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub


Public Sub CALCULAR(wsCosto As Currency)
Dim valor As Currency
Dim tC As Integer
Flag_Inicial = "A"
tC = grid_det.COL
valor = wsCosto * (1 + (val(grid_det.TextMatrix(grid_det.Row, 4)) / 100))
If valor < 0 Then
 ' grid_det.Col = 5
 ' grid_det.CellForeColor = QBColor(12)
Else
 ' grid_det.Col = 5
 ' grid_det.CellForeColor = QBColor(0)
End If
grid_det.TextMatrix(grid_det.Row, 5) = Format(valor, "0.0000") ' PRECIO 1
valor = wsCosto * (1 + (val(grid_det.TextMatrix(grid_det.Row, 6)) / 100))
If valor < 0 Then
 ' grid_det.Col = 7
 ' grid_det.CellForeColor = QBColor(12)
Else
 ' grid_det.Col = 7
 ' grid_det.CellForeColor = QBColor(0)
End If
grid_det.TextMatrix(grid_det.Row, 7) = Format(valor, "0.0000") ' PRECIO 2
valor = wsCosto * (1 + (val(grid_det.TextMatrix(grid_det.Row, 8)) / 100))
If valor < 0 Then
 ' grid_det.Col = 9
 ' grid_det.CellForeColor = QBColor(12)
Else
 ' grid_det.Col = 9
 ' grid_det.CellForeColor = QBColor(0)
End If
grid_det.TextMatrix(grid_det.Row, 9) = Format(valor, "0.0000") ' PRECIO 3
valor = wsCosto * (1 + (val(grid_det.TextMatrix(grid_det.Row, 10)) / 100))
If valor < 0 Then
 ' grid_det.Col = 11
 ' grid_det.CellForeColor = QBColor(12)
Else
 ' grid_det.Col = 11
 ' grid_det.CellForeColor = QBColor(0)
End If
grid_det.TextMatrix(grid_det.Row, 11) = Format(valor, "0.0000") ' PRECIO 4
valor = wsCosto * (1 + (val(grid_det.TextMatrix(grid_det.Row, 12)) / 100))
If valor < 0 Then
 ' grid_det.Col = 13
 ' grid_det.CellForeColor = QBColor(12)
Else
  'grid_det.Col = 13
  'grid_det.CellForeColor = QBColor(0)
End If
grid_det.TextMatrix(grid_det.Row, 13) = Format(valor, "0.0000") ' PRECIO 5
grid_det.COL = tC
Flag_Inicial = ""

End Sub
Public Sub CALCULAR_POR(WSPRE As Currency, WSCOL As Integer)
Dim valor As Currency
If val(grid_det.TextMatrix(grid_det.Row, 3)) <> 0 Then
  valor = (WSPRE * 100) / val(grid_det.TextMatrix(grid_det.Row, 2)) - 100
Else
  valor = 0
End If
Flag_Inicial = "A"
If valor < 0 Then
 ' grid_det.Col = WSCOL - 1
 ' grid_det.CellForeColor = QBColor(12)
Else
 ' grid_det.Col = WSCOL - 1
 ' grid_det.CellForeColor = QBColor(0)
End If
grid_det.COL = WSCOL
Flag_Inicial = ""

grid_det.TextMatrix(grid_det.Row, WSCOL - 1) = Format(valor, "0.0000")


End Sub

Public Sub CALCULAR_PRE(WSPOR As Currency, WSCOL As Integer)
Dim valor As Currency
valor = val(grid_det.TextMatrix(grid_det.Row, 2)) * (1 + (WSPOR / 100))
Flag_Inicial = "A"
If valor < 0 Then
 ' grid_det.Col = WSCOL + 1
 ' grid_det.CellForeColor = QBColor(12)
Else
 ' grid_det.Col = WSCOL + 1
 ' grid_det.CellForeColor = QBColor(0)
End If
grid_det.COL = WSCOL
Flag_Inicial = ""
grid_det.TextMatrix(grid_det.Row, WSCOL + 1) = Format(valor, "0.0000")

End Sub
Public Sub CAL()

If KeyAscii = 13 Then
  Flag_Change = ""
  If Flag_Change <> "A" And grid_det.Row = grid_det.Rows - 1 Then
   If grid_det.COL = 3 Then ' costo base
     If LK_FLAG_CALCULO = "A" Then
      CALCULAR val(grid_det.TextMatrix(grid_det.Row, 3))
     Else
      CALCULAR_OTRO val(grid_det.TextMatrix(grid_det.Row, 3))
     End If
   End If
   If grid_det.COL = 5 Or grid_det.COL = 7 Or grid_det.COL = 9 Or grid_det.COL = 11 Or grid_det.COL = 13 Then ' costo PORCENTAJE
    CALCULAR_POR val(grid_det.TextMatrix(grid_det.Row, grid_det.COL)), grid_det.COL
   End If
   If grid_det.COL = 4 Or grid_det.COL = 6 Or grid_det.COL = 8 Or grid_det.COL = 10 Or grid_det.COL = 12 Then ' costo PORCENTAJE
    CALCULAR_PRE val(grid_det.TextMatrix(grid_det.Row, grid_det.COL)), grid_det.COL
   End If
  Flag_Change = "A"
 End If
End If



End Sub
Public Sub COMPRAS_ULT(TOT_COMPRAS As Integer)
Dim wultimos As Integer
wultimos = TOT_COMPRAS
grid_compras.Rows = 1
grid_compras.Cols = 6
grid_compras.Clear
grid_compras.ColWidth(0) = 400
grid_compras.ColWidth(1) = 1000
grid_compras.ColWidth(2) = 2500
grid_compras.ColWidth(3) = 1200
grid_compras.ColWidth(4) = 900
grid_compras.ColWidth(5) = 900

grid_compras.TextMatrix(0, 0) = "Item"
grid_compras.TextMatrix(0, 1) = "Fecha"
grid_compras.TextMatrix(0, 2) = "Proveedor"
grid_compras.TextMatrix(0, 3) = "Unidad"
grid_compras.TextMatrix(0, 4) = "Cantidad"
grid_compras.TextMatrix(0, 5) = "Precios"
'wultimos = Val(txtlistar.text)
Set PSART_LOC = CN.CreateQuery("", pub_cadena)
'Set artloc_llave = PSART_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)

PSART_LOC.MaxRows = wultimos
pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ?  AND FAR_TIPMOV = ? AND FAR_CODART = ? ORDER BY FAR_FECHA"
Set PSART_LOC = CN.CreateQuery("", pub_cadena)
PSART_LOC(0) = 0
PSART_LOC(1) = 0
PSART_LOC(2) = 0
Set artloc_llave = PSART_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_LOC(0) = LK_CODCIA
PSART_LOC(1) = 20
PSART_LOC(2) = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 7))
artloc_llave.Requery
If artloc_llave.EOF Then
    MsgBox " No hay Compras . . .", 48, Pub_Titulo
    GoTo fin
End If
Fila = 0
grid_compras.Rows = 1
artloc_llave.MoveLast
Do Until artloc_llave.BOF
    Fila = Fila + 1
    grid_compras.Rows = grid_compras.Rows + 1
    grid_compras.TextMatrix(Fila, 0) = Fila
    grid_compras.TextMatrix(Fila, 1) = Format(artloc_llave!FAR_fecha_compra, "dd/mm/yyyy")
    SQ_OPER = 1
    pu_codclie = val(artloc_llave!far_codclie)
    pu_cp = "P"
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If cli_llave.EOF Then
        MsgBox "Error Grave  NO EXISTE Proveeedor ... ", 48, Pub_Titulo
        GoTo fin
    End If
    grid_compras.TextMatrix(Fila, 2) = Trim(cli_llave!CLI_NOMBRE)
    grid_compras.TextMatrix(Fila, 4) = Format(artloc_llave!far_cantidad / val(Grid_Arti.TextMatrix(Grid_Arti.Row, 8)), "0.0000")
    grid_compras.TextMatrix(Fila, 3) = Grid_Arti.TextMatrix(Grid_Arti.Row, 1)
    If artloc_llave!far_cantidad = 0 Then
      MsgBox "Cantidad en 0", , Pub_Titulo
      grid_compras.TextMatrix(Fila, 5) = "0"
    Else
      grid_compras.TextMatrix(Fila, 5) = Format(((artloc_llave!FAR_SUBTOTAL) - (artloc_llave!FAR_DESCTO)) / artloc_llave!far_cantidad) * (100 + LK_IGV) / 100 '* Val(Grid_Arti.TextMatrix(Grid_Arti.Row, 8), "0.00")
    End If
    If Fila = wultimos Then
     Exit Do
    End If
  artloc_llave.MovePrevious
Loop

Exit Sub
fin:
End Sub

Public Function COSTO_REPOS(WCODART As Currency)
Dim wultimos As Integer
Dim PSART_LOC4 As rdoQuery
Dim artloc_llave4 As rdoResultset
'PSART_LOC.MaxRows = 1
pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ?  AND FAR_TIPMOV = ? AND FAR_CODART = ? ORDER BY FAR_FECHA DESC"
Set PSART_LOC4 = CN.CreateQuery("", pub_cadena)
PSART_LOC4(0) = 0
PSART_LOC4(1) = 0
PSART_LOC4(2) = 0
Set artloc_llave4 = PSART_LOC4.OpenResultset(rdOpenKeyset, rdConcurValues)
PSART_LOC4(0) = LK_CODCIA
PSART_LOC4(1) = 20
PSART_LOC4(2) = WCODART
artloc_llave4.Requery
If artloc_llave4.EOF Then
    COSTO_REPOS = 0
    GoTo fin
End If
If artloc_llave4!far_cantidad = 0 Then
  COSTO_REPOS = 0
Else
  COSTO_REPOS = Format(artloc_llave4!far_precio_neto / artloc_llave4!far_cantidad, "0.0000")
End If

fin:
End Function

Public Function CONSIS_UNIDAD() As Boolean
Dim i As Integer
Dim QUIEN As String
QUIEN = 0
For Fila = 1 To grid_det.Rows - 1
   If Trim(grid_det.TextMatrix(Fila, 0)) <> "" And val(grid_det.TextMatrix(Fila, 1)) <> 0 Then
   Else
    MsgBox "Verificar Datos de unidad no valido..", 48, Pub_Titulo
    CONSIS_UNIDAD = False
    Exit Function
   End If
   If val(grid_det.TextMatrix(Fila, 4)) > 999.99 Or val(grid_det.TextMatrix(Fila, 6)) > 999.99 Or val(grid_det.TextMatrix(Fila, 8)) > 999.99 Or val(grid_det.TextMatrix(Fila, 10)) > 999.99 Or val(grid_det.TextMatrix(Fila, 12)) > 999.99 Then
     MsgBox " El Procentaje debe ser menor o igual a  999.99 .", 48, Pub_Titulo
     CONSIS_UNIDAD = False
     Exit Function
   End If
   If val(grid_det.TextMatrix(Fila, 4)) < -999.99 Or val(grid_det.TextMatrix(Fila, 6)) < -999.99 Or val(grid_det.TextMatrix(Fila, 8)) < -999.99 Or val(grid_det.TextMatrix(Fila, 10)) < -999.99 Or val(grid_det.TextMatrix(Fila, 12)) < -999.99 Then
     MsgBox " El Procentaje debe ser menor o igual a  -999.99 .", 48, Pub_Titulo
     CONSIS_UNIDAD = False
     Exit Function
   End If
Next Fila
CONSIS_UNIDAD = True
End Function

Public Sub GRABA_DATO(wSTOCK As Currency, w1Min_2Max As Integer)
Dim Wsec As Integer
Dim WCODPRO As Integer
Dim WCODART As Integer
Screen.MousePointer = 11
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_KEY = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 7))
LEER_ART_LLAVE
If art_LLAVE.EOF Then
  Screen.MousePointer = 0
  MsgBox "No actualizado ... Error Intente Nuevamente ...", 48, Pub_Titulo
  Exit Sub
End If
PRE_UNIDADES(0) = LK_CODCIA
PRE_UNIDADES(2) = "A"
PRE_UNIDADES(1) = PUB_KEY
preuni_llave.Requery
art_LLAVE.Edit
If w1Min_2Max = 1 Then
 art_LLAVE!art_stock_min = wSTOCK * preuni_llave!PRE_EQUIV
ElseIf w1Min_2Max = 2 Then
 art_LLAVE!ART_STOCK_MAX = wSTOCK * preuni_llave!PRE_EQUIV
End If
art_LLAVE.Update
Screen.MousePointer = 0
End Sub
Public Sub GRABA_POR(wPORCENTAJE As Currency, w12345 As Integer)
Dim Wsec As Integer
Dim WCODPRO As Integer
Dim WCODART As Integer
If w12345 = 99 Then
 SQ_OPER = 1
 pu_codcia = LK_CODCIA
 PUB_CODART = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 7))
 PUB_SECUEN = 0
 LEER_PRE_LLAVE
 If pre_llave.EOF Then
    Screen.MousePointer = 0
    MsgBox "Reiniciar Sistema..."
    Unload FrmPreUnidad
    Exit Sub
 End If
 pre_llave.Edit
 pre_llave!PRE_PRE1 = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 5))
 pre_llave!PRE_PRE2 = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 6))
 pre_llave.Update
 Screen.MousePointer = 0
 Exit Sub
End If

Screen.MousePointer = 11
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_KEY = val(Grid_Arti.TextMatrix(Grid_Arti.Row, 7))
LEER_ART_LLAVE
If art_LLAVE.EOF Then
  Screen.MousePointer = 0
  MsgBox "No actualizado ... Error Intente Nuevamente ...", 48, Pub_Titulo
  Exit Sub
End If
art_LLAVE.Edit
If w12345 = 1 Then
 art_LLAVE!ART_POR1 = wPORCENTAJE
ElseIf w12345 = 2 Then
 art_LLAVE!ART_POR2 = wPORCENTAJE
ElseIf w12345 = 3 Then
 art_LLAVE!ART_POR3 = wPORCENTAJE
ElseIf w12345 = 4 Then
 art_LLAVE!ART_POR4 = wPORCENTAJE
ElseIf w12345 = 5 Then
 art_LLAVE!ART_POR5 = wPORCENTAJE
End If
art_LLAVE.Update
Screen.MousePointer = 0
End Sub

Public Sub PRO_GOTFOCUS()
If Grid_Arti.Enabled = False Then
 Exit Sub
End If
grid_det.Cols = 1
grid_det.Rows = 1
grid_det.Clear
grid_det.Enabled = False
Grid_Arti.Clear
Grid_Arti.Rows = 1
Grid_Arti.Cols = 1
Grid_Arti.Enabled = False
grid_compras.Rows = 1
grid_compras.Cols = 1
grid_compras.Clear
grid_compras.Enabled = False
cmdcompras.Enabled = False
Txt_key.Text = ""

End Sub
Public Sub CALCULAR_OTRO(wsCosto As Currency)
Dim valor As Currency
Dim tC As Integer
If Flag_Inicial = "A" Then
  Exit Sub
End If
Flag_Inicial = "A"
tC = grid_det.COL
If wsCosto = 0 Then
  GoTo CERO
End If
valor = (val(grid_det.TextMatrix(grid_det.Row, 5) * 100)) / val(grid_det.TextMatrix(grid_det.Row, 3)) - 100
If val(grid_det.TextMatrix(grid_det.Row, 5)) = 0 Then
'   grid_det.TextMatrix(grid_det.Row, 4) = Format(0, "0.00") ' PRECIO 1
  'grid_det.Col = 4
  'grid_det.CellForeColor = QBColor(12)
Else
   grid_det.TextMatrix(grid_det.Row, 4) = Format(valor, "0.0000") ' PRECIO 1
  'grid_det.Col = 4
  'grid_det.CellForeColor = QBColor(0)
End If
valor = (val(grid_det.TextMatrix(grid_det.Row, 7) * 100)) / val(grid_det.TextMatrix(grid_det.Row, 3)) - 100
If val(grid_det.TextMatrix(grid_det.Row, 7)) = 0 Then
 '  grid_det.TextMatrix(grid_det.Row, 6) = Format(0, "0.00") ' PRECIO 1
  'grid_det.Col = 6
  'grid_det.CellForeColor = QBColor(12)
Else
   grid_det.TextMatrix(grid_det.Row, 6) = Format(valor, "0.0000") ' PRECIO 1
  'grid_det.Col = 6
  'grid_det.CellForeColor = QBColor(0)
End If
valor = (val(grid_det.TextMatrix(grid_det.Row, 9) * 100)) / val(grid_det.TextMatrix(grid_det.Row, 3)) - 100
If val(grid_det.TextMatrix(grid_det.Row, 9)) = 0 Then
  ' grid_det.TextMatrix(grid_det.Row, 8) = Format(0, "0.00") ' PRECIO 1
  'grid_det.Col = 8
  'grid_det.CellForeColor = QBColor(12)
Else
  grid_det.TextMatrix(grid_det.Row, 8) = Format(valor, "0.0000") ' PRECIO 1
  'grid_det.Col = 8
  'grid_det.CellForeColor = QBColor(0)
End If
valor = (val(grid_det.TextMatrix(grid_det.Row, 11) * 100)) / val(grid_det.TextMatrix(grid_det.Row, 3)) - 100
If val(grid_det.TextMatrix(grid_det.Row, 11)) = 0 Then
  'grid_det.Col = 10
  'grid_det.CellForeColor = QBColor(12)
  'grid_det.TextMatrix(grid_det.Row, 10) = Format(0, "0.00") ' PRECIO 1
Else
 ' grid_det.Col = 10
 ' grid_det.CellForeColor = QBColor(0)
 grid_det.TextMatrix(grid_det.Row, 10) = Format(valor, "0.0000") ' PRECIO 1
End If
valor = (val(grid_det.TextMatrix(grid_det.Row, 13) * 100)) / val(grid_det.TextMatrix(grid_det.Row, 3)) - 100
If val(grid_det.TextMatrix(grid_det.Row, 13)) = 0 Then
  'grid_det.Col = 12
  'grid_det.CellForeColor = QBColor(12)
  'grid_det.TextMatrix(grid_det.Row, 12) = Format(0, "0.00") ' PRECIO 1
Else
  'grid_det.Col = 12
 ' grid_det.CellForeColor = QBColor(0)
  grid_det.TextMatrix(grid_det.Row, 12) = Format(valor, "0.0000") ' PRECIO 1
End If
grid_det.COL = tC
Flag_Inicial = ""
Exit Sub
CERO:
   grid_det.TextMatrix(grid_det.Row, 5) = Format(0, "0.0000") ' PRECIO 1
   grid_det.TextMatrix(grid_det.Row, 7) = Format(0, "0.0000") ' PRECIO 2
   grid_det.TextMatrix(grid_det.Row, 9) = Format(0, "0.0000") ' PRECIO 3
   grid_det.TextMatrix(grid_det.Row, 11) = Format(0, "0.0000") ' PRECIO 4
   grid_det.TextMatrix(grid_det.Row, 13) = Format(0, "0.0000") ' PRECIO 5
  Flag_Inicial = ""
End Sub

Public Sub LLENADO_SUBFAM(wfami As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 123
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
    End If
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    listapro.ToolTipText = "TAB_TIPREG = 123"
    listapro.Clear
    Do Until tab_menor.EOF
        DoEvents
        listapro.AddItem tab_menor!tab_NOMLARGO & String(50, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        DoEvents
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
    
End Sub

Public Sub LLENA_COMBO(cont As ComboBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Public Sub FAMI_SUBFAM(wfami As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = 123
    PUB_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      PUB_CODCIA = "00"
    End If
    PUB_CODART = wfami
    SQ_OPER = 3
    LEER_TAB_LLAVE
    subfami.ToolTipText = "TAB_TIPREG = 123"
    subfami.Clear
    Do Until tab_menor.EOF
        DoEvents
        subfami.AddItem tab_menor!tab_NOMLARGO & String(90, " ") & Trim(CStr(tab_menor!TAB_NUMTAB))
        DoEvents
        CONTA = CONTA + 1
        tab_menor.MoveNext
    Loop
End Sub

Public Sub PROCESAR_DATA(ByVal sDescripccion As String, ByVal sMarca As String, ByVal sMedida As String, ByVal sCodigo As String, ByVal iIndex As Integer)
On Error GoTo Handler
Dim SQLSelect As String
Dim RSArticulos As rdoResultset
Dim PRMRSArticulos As rdoQuery
Dim K As Integer
Dim sHeader As String

    'sHeader = ">Nada|Codigo|Descripcion|Unidad|Stock|Marca||Medida"
    SQLSelect = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO, ARM_STOCK,  Art_Cuenta_Contab_c,Art_Cuenta_Contab, PRE_Unidad,pre_pre1,pre_pre2,pre_equiv,art_situacion "
    SQLSelect = SQLSelect & " FROM ARTI, ARTICULO, PRECIOS WHERE (ART_CODCIA = ARM_CODCIA) AND (PRE_CODART = ART_KEY) AND (PRE_CODCIA = ART_CODCIA) AND "
    SQLSelect = SQLSelect & "(ART_KEY = ARM_CODART) AND ART_KEY <> 0 AND ART_CODCIA = '" & LK_CODCIA & "' AND art_situacion <> 1 "
    'SQLSelect = SQLSelect & "AND TAB_CODCIA = '" & LK_CODCIA & "'  "
  
    If iIndex = 0 Then
        If Trim(sMedida) = "" Then
            SQLSelect = SQLSelect & "AND ART_NOMBRE LIKE '" & sDescripccion & "' AND ART_CUENTA_CONTAB_C Like '" & sMarca & "' and pre_secuencia = 0"
        Else
            SQLSelect = SQLSelect & "AND ART_NOMBRE LIKE '" & sDescripccion & "' AND ART_CUENTA_CONTAB_C Like '" & sMarca & "' AND Art_Cuenta_Contab = '" & sMedida & "' and pre_secuencia = 0"
        End If
        SQLSelect = SQLSelect & " ORDER BY ART_ALTERNO"
    ElseIf iIndex = 1 Then
        SQLSelect = SQLSelect + "AND ART_ALTERNO LIKE '" + sCodigo + "'"
        SQLSelect = SQLSelect & " ORDER BY ART_ALTERNO"
    End If
    'SETGRID sHeader, iIndex
    
    'RSArticulos.CursorLocation = adUseClient
    'RSArticulos.Open SQLSelect, StrConnAdo, adOpenForwardOnly, adLockOptimistic, adCmdText
    Set PRMRSArticulos = CN.CreateQuery("", SQLSelect)
    Set RSArticulos = PRMRSArticulos.OpenResultset(rdOpenKeyset, rdConcurValues)
    RSArticulos.Requery
    
    lblmensa.Caption = "Procesando . . . "
  DoEvents
  PB.max = RSArticulos.RowCount
  PB.Min = 0
  PB.Value = 0
  PB.Visible = True
  DoEvents
  Fila = 1
  Flag_Inicial = "A"
  Grid_Arti.Clear
  Grid_Arti.Rows = 3
  Grid_Arti.FixedRows = 2
  CABEZA_ARTI
  PRE_UNIDADES(0) = LK_CODCIA
  PRE_UNIDADES(2) = "A"
  WALT = String(8, " ")
  pu_codcia = LK_CODCIA
  SQ_OPER = 1
  Grid_Arti.Rows = IIf(RSArticulos.RowCount > 0, 1 + RSArticulos.RowCount, 2)
  Do While Not RSArticulos.EOF
  'Do Until artloc_llave.EOF
   PB.Value = PB.Value + 1
    PRE_UNIDADES(1) = Trim(RSArticulos("art_key"))
    preuni_llave.Requery
    PUB_CODART = Trim(RSArticulos("art_key"))
    LEER_ARM_LLAVE
    Fila = Fila + 1
    Grid_Arti.Rows = Fila + 1
    Grid_Arti.Row = Fila
    Grid_Arti.RowHeight(Fila) = 285
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      WALT = Left(RSArticulos("art_alterno"), 8)
      Grid_Arti.TextMatrix(Fila, 0) = "'" & WALT + " / " + Trim(RSArticulos("art_nombre")) + " / " + Trim(RSArticulos("art_cuenta_contab_c")) + " / " + Trim(RSArticulos("art_cuenta_contab"))
    Else
      WALT = Trim(RSArticulos("art_key"))
      Grid_Arti.TextMatrix(Fila, 0) = "'" & WALT + " / " + Trim(RSArticulos("art_nombre")) + " / " + Trim(RSArticulos("art_cuenta_contab_c")) + " / " + Trim(RSArticulos("art_cuenta_contab"))
    End If
    lblmensa.Caption = "Procesando . . .  " & Trim(RSArticulos("art_nombre"))
    DoEvents
    If preuni_llave.EOF Then
     'GoTo NOMUESTRA
    End If
    wsCostoRep = Nulo_Valor0(preuni_llave!PRE_COSTO) ' COSTO_REPOS(artloc_llave!art_KEY)
    Grid_Arti.TextMatrix(Fila, 1) = preuni_llave!pre_unidad
    Grid_Arti.TextMatrix(Fila, 2) = Format(arm_llave!arm_cospro * (100 + LK_IGV) / 100 * preuni_llave!PRE_EQUIV, "0.0000")
    Grid_Arti.TextMatrix(Fila, 3) = Format(preuni_llave!PRE_COSTO * (100 + LK_IGV) / 100 * preuni_llave!PRE_EQUIV, "0.0000") 'Format(wsCostoRep * preuni_llave!PRE_EQUIV, "0.00")
    Grid_Arti.TextMatrix(Fila, 4) = Format(preuni_llave!PRE_COSTO * 1.19, "0.0000")
    'If LK_EMP <> "PIU" Then
    Grid_Arti.TextMatrix(Fila, 5) = preuni_llave!PRE_PRE1
    Grid_Arti.TextMatrix(Fila, 6) = preuni_llave!PRE_PRE2
   ' Else
   '   Grid_Arti.TextMatrix(fila, 5) = Nulo_Valors(preuni_llave!PRE_PRE1)
   '   Grid_Arti.TextMatrix(fila, 6) = Nulo_Valors(preuni_llave!PRE_PRE2)
   ' End If
    
    Grid_Arti.TextMatrix(Fila, 7) = RSArticulos("art_key")
    Grid_Arti.TextMatrix(Fila, 8) = preuni_llave!PRE_EQUIV
    Grid_Arti.TextMatrix(Fila, 9) = Nulo_Valor0(arm_llave!arm_cospro)
    Grid_Arti.TextMatrix(Fila, 10) = wsCostoRep
   ' Grid_Arti.TextMatrix(fila, 11) = Format(Nulo_Valor0((RSArticulos("ART_STOCK_MIN"))) / preuni_llave!PRE_EQUIV, "0.00")
   ' Grid_Arti.TextMatrix(fila, 12) = Format(Nulo_Valor0(artloc_llave!ART_STOCK_MAX) / preuni_llave!PRE_EQUIV, "0.00")
   ' Grid_Arti.TextMatrix(fila, 13) = Nulo_Valor0(artloc_llave!ART_POR1)
   ' Grid_Arti.TextMatrix(fila, 14) = Nulo_Valor0(artloc_llave!ART_POR2)
   ' Grid_Arti.TextMatrix(fila, 15) = Nulo_Valor0(artloc_llave!ART_POR3)
   ' Grid_Arti.TextMatrix(fila, 16) = Nulo_Valor0(artloc_llave!ART_POR4)
   ' Grid_Arti.TextMatrix(fila, 17) = Nulo_Valor0(artloc_llave!ART_POR5)
   ' Grid_Arti.TextMatrix(fila, 19) = Nulo_Valor0(artloc_llave!ART_STOCK_MAX)
   ' Grid_Arti.TextMatrix(fila, 20) = Nulo_Valor0(artloc_llave!ART_STOCK_MIN)
    Grid_Arti.COL = 18
    Grid_Arti.CellForeColor = QBColor(2)
 
RSArticulos.MoveNext
'NOMUESTRA:
'   artloc_llave.MoveNext
  Loop
  PB.Visible = False
  lblmensa.Visible = False
  Grid_Arti.COL = 1
  Grid_Arti.Row = 2
  Flag_Inicial = ""
    
    RSArticulos.Close
    Set RSArticulos = Nothing
    Grid_Arti.SetFocus
    Grid_Arti.Row = 2
Exit Sub

'GoTo fin
'Exit Sub
'fin:
'fra2.Enabled = True
'PB.Visible = False
'lblmensa.Visible = False
'DoEvents
'CmdProcesa.Enabled = True
'DoEvents
'opciones.Enabled = True
 
    
    
    'grdarticulos.Rows = IIf(RSArticulos.RowCount > 0, 1 + RSArticulos.RowCount, 2)
    
    'Do While Not RSArticulos.EOF
    '    K = K + 1
    '    grdarticulos.TextMatrix(K, 1) = Trim((RSArticulos("Art_Alterno")))
    '    grdarticulos.TextMatrix(K, 2) = Trim(RSArticulos("Art_Nombre"))
    '    grdarticulos.TextMatrix(K, 3) = Left(IIf(IsNull(RSArticulos("pRE_Unidad")), "", RSArticulos("pRE_Unidad")), 8)
    '    grdarticulos.TextMatrix(K, 4) = Format(RSArticulos("ARM_STOCK") / RSArticulos("pre_equiv"), "0.00")
    '    grdarticulos.TextMatrix(K, 5) = Trim(RSArticulos("art_key"))
    '    grdarticulos.TextMatrix(K, 6) = RSArticulos("Art_cuenta_contab_c")
    '    grdarticulos.TextMatrix(K, 7) = RSArticulos("Art_cuenta_contab")
    '    grdarticulos.TextMatrix(K, 8) = Format(RSArticulos("pre_pre1"), "0.00")
    '    grdarticulos.TextMatrix(K, 9) = Format(RSArticulos("pre_pre2"), "0.00")
        'grdarticulos.TextMatrix(k, 7) = IIf(IsNull(RSArticulos("Art_Cuenta_Contab_C")), " ", RSArticulos("Art_Cuenta_Contab_C"))
       ' RSArticulos.MoveNext
    'Loop
    'grdarticulos_RowColChange
    RSArticulos.Close
    Set RSArticulos = Nothing
    Grid_Arti.SetFocus
    grdarticulos.Row = 1
    Exit Sub
Handler:
    MsgBox "Error en la Conexion"
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
Dim sCadena As String
Dim slDescripcion As String
Dim slMarca As String
Dim slMedida As String
Dim slCodigo As String
Dim ilPosSeparator As Integer
Dim ilTmpPosSep As Integer
    sCadena = txtdato.Text
    If KeyAscii = 13 And Len(sCadena) > 0 Then
        If sIndex = 0 Then
            ilPosSeparator = InStr(1, sCadena, ",", vbTextCompare)
            If ilPosSeparator = 0 Then
                MsgBox ("El dato ingresado no tiene la estructura de busqueda")
                txtdato.SetFocus
                Exit Sub
            End If
            slDescripcion = Mid(sCadena, 1, ilPosSeparator - 1) & "%"
            ilTmpPosSep = ilPosSeparator
            
            ilPosSeparator = InStr(ilPosSeparator + 1, sCadena, ",", vbTextCompare)
            If ilPosSeparator = 0 Then
                MsgBox ("El dato ingresado no tiene la estructura de busqueda")
                txtdato.SetFocus
                Exit Sub
            End If
            slMarca = Mid(sCadena, ilTmpPosSep + 1, ilPosSeparator - ilTmpPosSep - 1) & "%"
            
            slMedida = Mid(sCadena, ilPosSeparator + 1, Len(sCadena) - ilPosSeparator)
            If Trim(slDescripcion) = "" And (Trim(slMedida) = "" Or Trim(slMarca) = "") Then
                MsgBox "Error en los datos escritos para buscar"
            Exit Sub
        End If
        Else
            slCodigo = Trim(sCadena) & "%"
            If Trim(slCodigo) = "" Then
                MsgBox "Error en los datos escritos para buscar"
                Exit Sub
            End If
            
        End If
        'Grid_Arti.Clear
        PROCESAR_DATA slDescripcion, slMarca, slMedida, slCodigo, sIndex
    ElseIf KeyAscii = 27 Then
        frmbusqueda.Visible = False
        grdarticulos.Clear
        grdarticulos.Rows = 2
    End If

End Sub
Public Sub CALCULAR_PORCOL(ByVal WSPRE As Currency, ByVal WSCOL As Integer, ByVal WSROW As Integer)
Dim valor As Currency
If val(grid_det.TextMatrix(WSROW, 3)) <> 0 Then
  valor = (WSPRE * 100) / val(grid_det.TextMatrix(WSROW, 3)) - 100
Else
  valor = 0
End If
Flag_Inicial = "A"
If valor < 0 Then
  'grid_det.Col = WSCOL - 1
  'grid_det.CellForeColor = QBColor(12)
Else
  'grid_det.Col = WSCOL - 1
  'grid_det.CellForeColor = QBColor(0)
End If
grid_det.COL = WSCOL
Flag_Inicial = ""

grid_det.TextMatrix(WSROW, WSCOL - 1) = Format(valor, "0.0000")


End Sub
