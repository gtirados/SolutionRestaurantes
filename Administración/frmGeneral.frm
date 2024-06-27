VERSION 5.00
Begin VB.Form frmGeneral 
   Caption         =   "Datos Generales de la Empresa"
   ClientHeight    =   4995
   ClientLeft      =   1665
   ClientTop       =   1380
   ClientWidth     =   8880
   Icon            =   "frmGeneral.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2655
      Left            =   120
      TabIndex        =   32
      Top             =   720
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   600
         TabIndex        =   33
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame F1 
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   8655
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3720
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Gen_nombre"
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "Gen_Key"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Lbletiq 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Lbletiq 
         Caption         =   "LLave :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame F3 
      Caption         =   "Otros"
      Height          =   3255
      Left            =   6120
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtTiempo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Top             =   2640
         Width           =   615
      End
      Begin VB.CheckBox chkStock 
         Caption         =   "Controla Stock de Insumos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   36
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   1320
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   12
         Text            =   "gen_fecha_ref"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   11
         Text            =   "gen_fecha_ref"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   10
         Text            =   "gen_fecha_ref"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   9
         Text            =   "gen_fecha_ref"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "gen_dias_nolab"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox gen_bloqueo 
         Caption         =   " Bloquear Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "seg."
         Height          =   195
         Left            =   2040
         TabIndex        =   39
         Top             =   2760
         Width           =   300
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo:"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   37
         Top             =   2760
         Width           =   570
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "PWD :"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   " UID :"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "Name Server"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   930
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "G. fecha_ref :"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "G. dias_nolab :"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   1680
      TabIndex        =   14
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Frame F2 
      Caption         =   "Tasas "
      Height          =   3255
      Left            =   2880
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txticbper 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   40
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   1560
         TabIndex        =   34
         Text            =   "gen_tipo_cambio"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1560
         TabIndex        =   5
         Text            =   "gen_tasa_descto2"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   1560
         TabIndex        =   4
         Text            =   "gen_tasa_descto1"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1560
         TabIndex        =   3
         Text            =   "gen_igv"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   7
         Text            =   "gen_tasa_leg_adel"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1560
         TabIndex        =   6
         Text            =   "gen_tasa_leg_venc"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TxtGen 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Text            =   "gen_tasa_venta"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblICBPER 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ICBPER"
         Height          =   195
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cambio:"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Descuento 2 :"
         Height          =   195
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1410
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Descuento 1 :"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1410
      End
      Begin VB.Label Lbletiq 
         Caption         =   "I.G.V."
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "G. tasa_leg_adel :"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   1290
      End
      Begin VB.Label Lbletiq 
         AutoSize        =   -1  'True
         Caption         =   "G. tasa_leg_venc :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   1350
      End
      Begin VB.Label Lbletiq 
         Caption         =   "Tasa de Venta"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim PSpre_todo As rdoQuery
Dim pre_todo  As rdoResultset


Public Sub LIMPIA_GEN()
Dim i As Integer
For i = 0 To 13
    TxtGen(i).text = ""
Next i
'gen_bloqueo.Value = 0

End Sub


Private Sub cmdcerrar_Click()
Dim i
If Left(CmdCerrar.Caption, 3) = "&Ce" Then
    'SALIMOS
     Unload frmGeneral
Else
    '* CANCELAMOS
     Screen.MousePointer = 11
     CmdModificar.Caption = "&Modificar"
     CmdCerrar.Caption = "&Cerrar"
     For i = 0 To 13
        TxtGen(i).Enabled = False
     Next i
     Me.txtTiempo.Enabled = False
'     gen_bloqueo.Enabled = False

     Screen.MousePointer = 0
End If
End Sub

Private Sub CmdModificar_Click()
Dim i, respuesta, estilo, pub_mensaje
If Left(CmdModificar.Caption, 2) = "&M" Then
    For i = 0 To 13
        TxtGen(i).Enabled = True
    Next i
    Me.txtTiempo.Enabled = True
'    gen_bloqueo.Enabled = True
    CmdModificar.Caption = "&Grabar"
    CmdCerrar.Caption = "&Cancelar"
Else
    pub_mensaje = " ¿Desea Modificar los Datos... Usuario ADMIN ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
       Exit Sub
    End If

    Screen.MousePointer = 11
    PS_GEN(0) = 0
    gen.Requery
    gen.Edit
    gen.rdoColumns(0) = Nulo_Valor0(TxtGen(0).text)
    gen.rdoColumns(1) = Nulo_Valors(TxtGen(1).text)
    gen.rdoColumns(2) = Nulo_Valor0(TxtGen(2).text)
    gen.rdoColumns(3) = Nulo_Valor0(TxtGen(3).text)
    gen.rdoColumns(4) = Nulo_Valor0(TxtGen(4).text)
    gen.rdoColumns(5) = Nulo_Valor0(TxtGen(5).text)
    gen.rdoColumns(6) = Nulo_Valor0(TxtGen(6).text)
    gen.rdoColumns(7) = Nulo_Valor0(TxtGen(7).text)
    gen.rdoColumns(8) = Nulo_Valor0(TxtGen(8).text)
    gen.rdoColumns(9) = Nulo_Valors(TxtGen(9).text)
    gen.rdoColumns(11) = Nulo_Valors(TxtGen(10).text)
    gen.rdoColumns(12) = Nulo_Valors(TxtGen(11).text)
    gen.rdoColumns(13) = Nulo_Valors(TxtGen(12).text)
    gen.rdoColumns!GEN_TIEMPO_ELIMINAPLATO = Me.txtTiempo.text
    gen.rdoColumns!GEN_ICBPER = Me.txticbper.text
    gen.rdoColumns!gen_tipo_cambio = Nulo_Valor0(TxtGen(13).text)
    'gen.rdoColumns(10) = gen_bloqueo.Value
    Pub_Titulo = Trim(gen!GEN_NOMBRE)
    PUB_RUTA_REPORTE = Trim(Nulo_Valors(gen!gen_ruta_reportes))
    PUB_RUTA_OTRO = Trim(Nulo_Valors(gen!gen_ruta_otros))
    LK_IGV = Nulo_Valor0(gen!GEN_IGV)
    LK_TIPO_CAMBIO = Nulo_Valor0(gen!gen_tipo_cambio)
    If chkStock.Value = 1 Then
     gen.rdoColumns!gen_controla_stock = "A"
    Else
     gen.rdoColumns!gen_controla_stock = " "
    End If
    
    gen.Update
    CmdModificar.Caption = "&Modificar"
    CmdCerrar.Caption = "&Cerrar"
    Screen.MousePointer = 0
    For i = 0 To 13
        TxtGen(i).Enabled = False
    Next i
'    gen_bloqueo.Enabled = False

    MsgBox "OK. Modificado ... ", 48, Pub_Titulo
End If

End Sub

Private Sub Form_Activate()
Dim i
'llena los datos
PS_GEN(0) = 0
gen.Requery
For i = 0 To 9
    TxtGen(i).text = gen.rdoColumns(i)
Next i
'gen_bloqueo.Value = gen.rdoColumns(10)
TxtGen(10).text = Nulo_Valors(gen.rdoColumns(11))
TxtGen(11).text = Nulo_Valors(gen.rdoColumns(12))
TxtGen(12).text = Nulo_Valors(gen.rdoColumns(13))
TxtGen(13).text = Nulo_Valors(gen.rdoColumns!gen_tipo_cambio)
            Me.txtTiempo.text = gen.rdoColumns!GEN_TIEMPO_ELIMINAPLATO
            Me.txticbper.text = gen.rdoColumns!GEN_ICBPER
F1.Visible = True
F2.Visible = True
F3.Visible = True
For i = 0 To 13
    TxtGen(i).Enabled = False
Next i
Me.txtTiempo.Enabled = False
'gen_bloqueo.Enabled = False
chkStock.Value = 0
If Nulo_Valors(gen.rdoColumns!gen_controla_stock) = "A" Then
   chkStock.Value = 1
End If


End Sub

Private Sub Form_Load()
CenterMe frmGeneral
LIMPIA_GEN
End Sub

