VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Begin VB.Form Frmislas 
   Caption         =   "Islas "
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   ControlBox      =   0   'False
   Icon            =   "Frmislas.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Frame frapro 
      Caption         =   "Producto:"
      Height          =   735
      Left            =   2400
      TabIndex        =   48
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txt_arti 
         Height          =   285
         Left            =   600
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   495
      Left            =   6480
      TabIndex        =   50
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Timer PARPADEA 
      Interval        =   200
      Left            =   120
      Top             =   4560
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   4080
      TabIndex        =   37
      Top             =   4560
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame F1 
      Caption         =   "Isla :"
      Height          =   2415
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   9375
      Begin VB.TextBox textovar 
         Height          =   285
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   56
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ListBox productos 
         Height          =   645
         Left            =   2520
         TabIndex        =   2
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid TURNOS 
         Height          =   1095
         Left            =   6240
         TabIndex        =   3
         ToolTipText     =   "Click para seleccionar el responsable"
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1931
         _Version        =   327680
         FixedRows       =   0
      End
      Begin VB.ComboBox lisresp 
         Height          =   315
         Left            =   6360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtnombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         MaxLength       =   30
         TabIndex        =   1
         Top             =   120
         Width           =   4815
      End
      Begin VB.TextBox Txt_key 
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
         Height          =   285
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid gridpro 
         Height          =   1575
         Left            =   120
         TabIndex        =   53
         ToolTipText     =   "Click para seleccionar el responsable"
         Top             =   720
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2778
         _Version        =   327680
         FixedCols       =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Turnos:"
         Height          =   195
         Index           =   19
         Left            =   5520
         TabIndex        =   51
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "[INSERT]=Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   12
         Left            =   4080
         TabIndex        =   47
         Top             =   550
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "[INSERT]=Mantenimiento"
         Height          =   195
         Index           =   11
         Left            =   6600
         TabIndex        =   46
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Lista de Surtidores:"
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Responsable"
         Height          =   255
         Index           =   13
         Left            =   6960
         TabIndex        =   41
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   28
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   27
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame F2 
      Height          =   2055
      Left            =   0
      TabIndex        =   29
      Top             =   2400
      Visible         =   0   'False
      Width           =   9375
      Begin VB.TextBox remi 
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
         Height          =   285
         Left            =   8400
         MaxLength       =   4
         TabIndex        =   54
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox numfac_p 
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
         Left            =   600
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1100
         Width           =   975
      End
      Begin VB.TextBox serie_p 
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
         Left            =   600
         MaxLength       =   4
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox numfac_p_f 
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicializar - Serie Ped"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox chefactura 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicializar - Serie Factura"
         Height          =   255
         Left            =   6840
         TabIndex        =   17
         Top             =   120
         Width           =   2055
      End
      Begin VB.CheckBox cheboleta 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicializar - Serie Boleta "
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox cheguia 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicializar - Serie Guia "
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox numfac_f_f 
         Height          =   285
         Left            =   6960
         TabIndex        =   20
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox numfac_b_f 
         Height          =   285
         Left            =   5280
         TabIndex        =   16
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox numfac_g_f 
         Height          =   285
         Left            =   2640
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox numfac_f 
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
         Height          =   285
         Left            =   6960
         MaxLength       =   9
         TabIndex        =   19
         Top             =   1100
         Width           =   975
      End
      Begin VB.TextBox serie_f 
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
         Height          =   285
         Left            =   6960
         MaxLength       =   4
         TabIndex        =   18
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox serie_g 
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox numfac_g 
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
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   11
         Top             =   1100
         Width           =   975
      End
      Begin VB.TextBox Serie_b 
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
         Height          =   285
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox numfac_b 
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
         Height          =   285
         Left            =   5280
         MaxLength       =   9
         TabIndex        =   15
         Top             =   1100
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Serie de Guía de Remisión :"
         Height          =   420
         Index           =   18
         Left            =   8280
         TabIndex        =   55
         Top             =   480
         Width           =   1050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Serie Pedido :"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Ped. Inicial  :"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   44
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Ped. Final :"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   43
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Factura Final :"
         Height          =   255
         Index           =   10
         Left            =   6840
         TabIndex        =   40
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Boleta Final :"
         Height          =   255
         Index           =   9
         Left            =   4440
         TabIndex        =   39
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Guia Final :"
         Height          =   255
         Index           =   8
         Left            =   2280
         TabIndex        =   38
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Factura Inicial:"
         Height          =   375
         Index           =   7
         Left            =   6840
         TabIndex        =   36
         Top             =   900
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Serie Factura :"
         Height          =   255
         Index           =   3
         Left            =   6840
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Guia Inicial  :"
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   34
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Serie Guia :"
         Height          =   255
         Index           =   5
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   855
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Serie Boleta :"
         Height          =   255
         Index           =   4
         Left            =   4440
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "N.Fac Boleta Inicial:"
         Height          =   255
         Index           =   2
         Left            =   4320
         TabIndex        =   30
         Top             =   900
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   3840
      TabIndex        =   23
      Top             =   4800
      Width           =   1485
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   7320
      TabIndex        =   25
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "La definicón de los documentos está configurada por compañia."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   600
      TabIndex        =   52
      Top             =   3240
      Width           =   6750
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   32
      Top             =   4560
      Width           =   900
   End
End
Attribute VB_Name = "Frmislas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WCOD_ORIGINAL As Currency
Dim VAR_ACTIVAR As Integer
Dim PSVE2_LLAVE  As rdoQuery
Dim VE2_LLAVE As rdoResultset
Dim pasa As Boolean
Dim loc_key As Integer
Dim CU As Integer
Dim flag_tempo As String * 1
Dim PSMA_LLAVE As rdoQuery
Dim man_llave As rdoResultset
Dim man_busca  As rdoResultset
Dim PSMA_BUSCA As rdoQuery


Public Function GENERA_VEN() As Integer
Dim valor As Integer
Dim ven_loc As rdoResultset
Dim PSVEN_LOC  As rdoQuery
pub_cadena = "SELECT VEM_CODVEN FROM VEMAEST WHERE VEM_CODCIA  = ?  ORDER BY VEM_CODVEN"
Set PSVEN_LOC = CN.CreateQuery("", pub_cadena)
Set ven_loc = PSVEN_LOC.OpenResultset(rdOpenKeyset, rdConcurValues)
PSVEN_LOC(0) = LK_CODCIA
ven_loc.Requery
If ven_loc.EOF Then
 valor = 0
Else
 ven_loc.MoveLast
 valor = ven_loc!VEM_CODVEN
End If
GENERA_VEN = valor + 1

End Function

Public Sub GRABAR_VEN()
Dim WTAB_COD As Currency
Dim WISLA As Integer
If Left(cmdmodificar.Caption, 2) = "&G" Then
   ven_llave.Edit
Else
   ven_llave.AddNew
End If
ven_llave!VEM_CODVEN = Val(Frmislas.txt_key.Text)
ven_llave!vem_nombre = Frmislas.txtnombre.Text
ven_llave!vem_codcia = LK_CODCIA
ven_llave!VEM_SERIE_G = Val(Frmislas.serie_g.Text)
ven_llave!VEM_NUMFAC_G_INI = Val(Frmislas.numfac_g.Text)
ven_llave!VEM_SERIE_B = Val(Frmislas.Serie_b.Text)
ven_llave!VEM_NUMFAC_B_INI = Val(Frmislas.numfac_b.Text)
ven_llave!VEM_SERIE_F = Val(Frmislas.serie_f.Text)
ven_llave!VEM_NUMFAC_F_INI = Val(Frmislas.numfac_f.Text)
ven_llave!VEM_NUMFAC_G_FIN = Val(Frmislas.numfac_g_f.Text)
ven_llave!VEM_NUMFAC_B_FIN = Val(Frmislas.numfac_b_f.Text)
ven_llave!VEM_NUMFAC_F_FIN = Val(Frmislas.numfac_f_f.Text)
ven_llave!VEM_SERIE_R = Val(Frmislas.remi.Text)
ven_llave!VEM_SERIE_P = Val(Frmislas.serie_p.Text)
ven_llave!VEM_NUMFAC_P_INI = Val(Frmislas.numfac_p.Text)
ven_llave!VEM_NUMFAC_P_FIN = Val(Frmislas.numfac_p_f.Text)
ven_llave!VEM_FLAG_P = " "
If Check1.Value = 1 Then
  ven_llave!VEM_FLAG_P = "A"
End If


ven_llave!VEM_VENDEDOR = Val(Right(lisresp.Text, 4))
ven_llave!VEM_FLAG_G = " "
ven_llave!VEM_FLAG_B = " "
ven_llave!VEM_FLAG_F = " "
If cheguia.Value = 1 Then
  ven_llave!VEM_FLAG_G = "A"
End If
If cheboleta.Value = 1 Then
  ven_llave!VEM_FLAG_B = "A"
End If
If chefactura.Value = 1 Then
  ven_llave!VEM_FLAG_F = "A"
End If
ven_llave.Update

For fila = 0 To TURNOS.Rows - 1
 PSVE2_LLAVE(0) = LK_CODCIA
 PSVE2_LLAVE(1) = txt_key.Text
 PSVE2_LLAVE(2) = Val(TURNOS.TextMatrix(fila, 0))
 VE2_LLAVE.Requery
 If VE2_LLAVE.EOF Then
   VE2_LLAVE.AddNew
   VE2_LLAVE!VE2_CODCIA = LK_CODCIA
   VE2_LLAVE!VE2_ISLA = txt_key.Text
   VE2_LLAVE!VE2_TURNO = Val(TURNOS.TextMatrix(fila, 0))
   VE2_LLAVE!VE2_CODRES = Val(TURNOS.TextMatrix(fila, 2))
 Else
   VE2_LLAVE.Edit
   VE2_LLAVE!VE2_CODRES = Val(TURNOS.TextMatrix(fila, 2))
 End If
 VE2_LLAVE.Update
Next fila



WTAB_COD = 2105
WISLA = Val(Frmislas.txt_key.Text)
' actualiza tabla de productos
pub_cadena = "DELETE TABLAS WHERE TAB_CONTABLE2 = " & WISLA & " AND TAB_TIPREG = " & WTAB_COD & " AND TAB_CODCIA  = '" & LK_CODCIA & "'"
CN.Execute pub_cadena, rdExecDirect
For fila = 1 To gridpro.Rows - 1
  tab_llave.AddNew
  tab_llave!TAB_CODCIA = LK_CODCIA
  tab_llave!TAB_TIPREG = WTAB_COD
  tab_llave!tab_numtab = Val(gridpro.TextMatrix(fila, 0))
  tab_llave!tab_nomlargo = Trim(gridpro.TextMatrix(fila, 5))
  tab_llave!tab_nomcorto = ""
  tab_llave!tab_codART = Val(gridpro.TextMatrix(fila, 3))
  tab_llave!TAB_contable2 = WISLA
  tab_llave.Update
Next fila


End Sub
Public Sub MENSAJE_VEN(TEXTO As String)
  LblMensaje.Caption = TEXTO
  PARPADEA.Enabled = True
End Sub

Public Sub LLENA_VEN(ban As Integer)
Dim i As Integer
If ban = 0 Then
       If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
         Else
          txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
       End If
       PUB_CODVEN = Val(txt_key.Text)
       pu_codcia = LK_CODCIA
       PUB_CODCIA = LK_CODCIA
       SQ_OPER = 1
       LEER_VEN_LLAVE
End If

Frmislas.txt_key.Text = Trim(Nulo_Valors(ven_llave!VEM_CODVEN))
Frmislas.txtnombre.Text = Trim(Nulo_Valors(ven_llave!vem_nombre))
Frmislas.serie_g.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_G))
Frmislas.numfac_g.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_G_INI))
Frmislas.Serie_b.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_B))
Frmislas.serie_p.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_P))
Frmislas.numfac_b.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_B_INI))
Frmislas.serie_f.Text = Trim(Nulo_Valors(ven_llave!VEM_SERIE_F))
Frmislas.numfac_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_F_INI))
Frmislas.numfac_p.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_P_INI))
Frmislas.numfac_g_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_G_FIN))
Frmislas.numfac_b_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_B_FIN))
Frmislas.numfac_f_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_F_FIN))
Frmislas.numfac_p_f.Text = Trim(Nulo_Valors(ven_llave!VEM_NUMFAC_P_FIN))
Frmislas.remi.Text = Nulo_Valor0(ven_llave!VEM_SERIE_R)

If Not IsNull(ven_llave!VEM_FECHA_ING) Then
  'txtfechaing.Text = Format(Nulo_Valors(ven_llave!VEM_FECHA_ING), "dd/mm/yyyy")
End If

PSMA_LLAVE(0) = LK_CODCIA
PSMA_LLAVE(1) = 2105
PSMA_LLAVE(2) = Val(Frmislas.txt_key.Text)
man_llave.Requery
cabe_pro
fila = 1
SQ_OPER = 1
pu_codcia = LK_CODCIA
If man_llave.EOF Then
  gridpro.Rows = gridpro.Rows + 1
  gridpro.RowHeight(fila) = 285
  gridpro.COL = 2
Else
Do Until man_llave.EOF
    PUB_KEY = man_llave!tab_codART
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then GoTo OTRO
    gridpro.Rows = gridpro.Rows + 1
    gridpro.RowHeight(fila) = 285
    gridpro.TextMatrix(fila, 0) = man_llave!tab_numtab
    gridpro.TextMatrix(fila, 1) = art_LLAVE!art_alterno
    gridpro.TextMatrix(fila, 2) = art_LLAVE!art_nombre
    gridpro.TextMatrix(fila, 3) = man_llave!tab_codART
    gridpro.TextMatrix(fila, 4) = man_llave!TAB_contable2
    gridpro.TextMatrix(fila, 5) = Format(man_llave!tab_nomlargo, "0.00")
    fila = fila + 1
OTRO:
    man_llave.MoveNext
Loop
End If
'ASIGNA_INT lisresp, Nulo_Valor0(ven_llave!VEM_VENDEDOR)

'txtfechaing.Mask = "##/##/####"
'frmislas.txtdireccion.Text = Trim(Nulo_Valors(ven_llave!VEM_DIRECCION))
'frmislas.txttelecasa.Text = Trim(Nulo_Valors(ven_llave!VEM_TELE_CASA))
'frmislas.txttelecelu.Text = Trim(Nulo_Valors(ven_llave!VEM_TELE_CELU))

cheguia.Value = 0
cheboleta.Value = 0
chefactura.Value = 0
Check1.Value = 0
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_G)) = "A" Then
  cheguia.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_B)) = "A" Then
  cheboleta.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_F)) = "A" Then
  chefactura.Value = 1
End If
If UCase(Nulo_Valors(ven_llave!VEM_FLAG_P)) = "A" Then
  Check1.Value = 1
End If
For fila = 0 To TURNOS.Rows - 1
 PSVE2_LLAVE(0) = LK_CODCIA
 PSVE2_LLAVE(1) = txt_key.Text
 PSVE2_LLAVE(2) = Val(TURNOS.TextMatrix(fila, 0))
 VE2_LLAVE.Requery
 If Not VE2_LLAVE.EOF Then TURNOS.TextMatrix(fila, 2) = VE2_LLAVE!VE2_CODRES
Next fila
If TURNOS.Rows > 0 Then
  TURNOS.COL = 1
  TURNOS.Row = 0
  TURNOS_EnterCell
End If
PUB_TIPREG = Val(Frmislas.txt_key.Text) + 3000
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
  productos.AddItem tab_mayor!tab_nomlargo + String(80, " ") + Str(tab_mayor!tab_codART)
  tab_mayor.MoveNext
Loop



End Sub
Public Sub LIMPIA_VEN()
txt_key.Text = ""
txtnombre.Text = ""
serie_g.Text = ""
numfac_g.Text = ""
Serie_b.Text = ""
numfac_b.Text = ""
serie_f.Text = ""
numfac_f.Text = ""
numfac_g_f.Text = ""
numfac_b_f.Text = ""
numfac_f_f.Text = ""
cheguia.Value = 0
cheboleta.Value = 0
chefactura.Value = 0
Check1.Value = 0
serie_p.Text = ""
numfac_p.Text = ""
numfac_p_f.Text = ""
lisresp.ListIndex = -1
productos.Clear
remi.Text = ""
For fila = 0 To TURNOS.Rows - 1
  TURNOS.TextMatrix(fila, 2) = 0
Next fila
cabe_pro


End Sub

Private Sub cheboleta_Click()
If Serie_b.Enabled Then
 Serie_b.SetFocus
End If
End Sub

Private Sub chefactura_Click()
If serie_f.Enabled Then
 serie_f.SetFocus
End If
End Sub

Private Sub cheguia_Click()
If serie_g.Enabled Then
 serie_g.SetFocus
End If
End Sub

Private Sub cmdagregar_Click()
'On Error GoTo ESCAPA
Dim wm As String
If Left(CmdAgregar.Caption, 2) = "&A" Then
    CmdAgregar.Caption = "&Grabar"
    cmdcancelar.Enabled = True
    cmdmodificar.Enabled = False
    cmdeliminar.Enabled = False
    LIMPIA_VEN
    DESBLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
    DESBLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, Check1, productos, lisresp, TURNOS
    remi.Enabled = True
    Frmislas.txt_key = GENERA_VEN
    Frmislas.txtnombre.SetFocus
    'AGREGAMOS EN BLANCO
Else
   If Frmislas.txtnombre.Text = "" Or Len(Frmislas.txtnombre.Text) = 0 Then
       MsgBox "Ingrese Nombre de Vendedor ..!!!", 48, Pub_Titulo
       Azul txtnombre, txtnombre
       Exit Sub
   End If
   wm = CONS_SUR
   If Trim(wm) <> "" Then
     MsgBox wm, 48, Pub_Titulo
     Exit Sub
   End If
   
   '"SI GRABA.."
    SQ_OPER = 1
    PUB_CODVEN = Val(Frmislas.txt_key.Text)
    pu_codcia = LK_CODCIA
    LEER_VEN_LLAVE
    If Not ven_llave.EOF Then
       MsgBox "Registro ,  EXISTE ... ", 48, Pub_Titulo
       Azul Frmislas.txt_key, txt_key
       Exit Sub
    End If
   Screen.MousePointer = 11
   GRABAR_VEN
   MENSAJE_VEN "Bancos , AGREGADO... "
   CmdAgregar.Caption = "&Agregar"
   cmdeliminar.Enabled = True
   cmdmodificar.Enabled = True
   LIMPIA_VEN
   BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
   BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, Check1, productos, lisresp, TURNOS
   remi.Enabled = False
   txt_key.Locked = False
   txt_key.SetFocus
   Screen.MousePointer = 0
      
End If
   
End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    
End If

End Sub

Private Sub cmdcancelar_Click()
textovar.Visible = False
If Left(CmdAgregar.Caption, 2) = "&A" And Left(cmdmodificar.Caption, 2) = "&M" Then
    LIMPIA_VEN
    txt_key.Locked = False
    MENSAJE_VEN "Proceso Cancelado... !!!    "
    txt_key.Enabled = True
    txt_key.SetFocus
     Exit Sub
End If
     Screen.MousePointer = 11
     If Left(cmdmodificar.Caption, 2) = "&G" Then
        cmdmodificar.Caption = "&Modificar"
        LLENA_VEN 1
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, Check1, productos, lisresp, TURNOS
        remi.Enabled = False
        txt_key.Locked = True
     Else
        CmdAgregar.Caption = "&Agregar"
        LIMPIA_VEN
        BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
        BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, Check1, productos, lisresp, TURNOS
        remi.Enabled = False
        txt_key.Locked = False
     End If
     cmdCerrar.Caption = "Ce&rrar"
     cmdcancelar.Enabled = True
     CmdAgregar.Enabled = True
     cmdmodificar.Enabled = True
     cmdeliminar.Enabled = True
     txt_key.Enabled = True
     MENSAJE_VEN "Proceso Cancelado... !!!    "
     txt_key.SetFocus
     Screen.MousePointer = 0

End Sub

Private Sub cmdCerrar_Click()
ws_conta = 0
Unload Frmislas
End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Frmislas.txt_key.SetFocus
End If

End Sub

Private Sub cmdeliminar_Click()
Dim WTAB_COD  As Currency
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

If Len(txt_key) = 0 Or Len(txtnombre) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
  pub_cadena = "SELECT FAR_CODVEN FROM FACART WHERE FAR_CODCIA = ? AND FAR_ISLA = ? "
  Set PS_REP01 = CN.CreateQuery("", pub_cadena)
  PS_REP01.MaxRows = 1
  Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  PS_REP01(0) = LK_CODCIA
  PS_REP01(1) = ven_llave!VEM_CODVEN
  llave_rep01.Requery
  If Not llave_rep01.EOF Then
     Screen.MousePointer = 0
     MsgBox "NO se Puede Eliminar ...  Vendedor  TIENE H I S T O R I A.. ", 48, Pub_Titulo
     Exit Sub
  End If
  
  pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    Screen.MousePointer = 11
    ven_llave.Delete
    WTAB_COD = Val(txt_key.Text)
    pub_cadena = "DELETE TABLAS WHERE TAB_CONTABLE2 = " & WTAB_COD & " AND TAB_TIPREG = " & 2105 & " AND TAB_CODCIA  = '" & LK_CODCIA & "'"
    CN.Execute pub_cadena, rdExecDirect
    txt_key.Text = ""
    txt_key.Locked = False
    LIMPIA_VEN
    MENSAJE_VEN "Registro   ELIMINADO ... "
    Screen.MousePointer = 0
   Exit Sub
  End If
  
  Screen.MousePointer = 0
End Sub

Private Sub cmdmodificar_Click()
Dim wm     As String

If Len(txt_key) = 0 Then
   MENSAJE_VEN "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdmodificar.Caption, 2) = "&M" Then
    cmdmodificar.Caption = "&Grabar"
    CmdAgregar.Enabled = False
    cmdeliminar.Enabled = False
    cmdcancelar.Enabled = True
    txt_key.Locked = True
    DESBLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
    DESBLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, Check1, productos, lisresp, TURNOS
    remi.Enabled = True
    txtnombre.SetFocus
Else
    '*Grabar las modificaciones
    If txtnombre.Text = "" Or Len(txtnombre.Text) = 0 Then
         MsgBox " Nombre Invalido ....", 48, Pub_Titulo
         Exit Sub
    End If
    wm = CONS_SUR
    If Trim(wm) <> "" Then
     MsgBox wm, 48, Pub_Titulo
     gridpro.SetFocus
     Exit Sub
    End If
    
     Screen.MousePointer = 11
     GRABAR_VEN
     MENSAJE_VEN "Registro , MODIFICADO... "
     cmdmodificar.Caption = "&Modificar"
     cmdcancelar.Enabled = True
     CmdAgregar.Enabled = True
     cmdeliminar.Enabled = True
     txt_key.Locked = True
     BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
     BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, Check1, productos, lisresp, TURNOS
     remi.Enabled = False
     Screen.MousePointer = 0
End If

End Sub


Private Sub Form_Load()
Unload FORM_GRIFO
pub_cadena = "SELECT * FROM VEMAEST2 WHERE VE2_CODCIA  = ? AND VE2_ISLA  = ?  AND VE2_TURNO = ? ORDER BY VE2_CODCIA "
Set PSVE2_LLAVE = CN.CreateQuery("", pub_cadena)
PSVE2_LLAVE(0) = 0
PSVE2_LLAVE(1) = 0
PSVE2_LLAVE(2) = 0
Set VE2_LLAVE = PSVE2_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PUB_TIPREG = 2103
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
fila = 0
Do Until tab_mayor.EOF
  lisresp.AddItem tab_mayor!tab_nomlargo + String(80, " ") + Format(tab_mayor!tab_numtab, "000")
  tab_mayor.MoveNext
Loop

pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  AND TAB_CONTABLE2 = ? ORDER BY TAB_NUMTAB "
Set PSMA_LLAVE = CN.CreateQuery("", pub_cadena)
PSMA_LLAVE(0) = 0
PSMA_LLAVE(1) = 0
PSMA_LLAVE(2) = 0
Set man_llave = PSMA_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT TAB_CONTABLE2 FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  AND TAB_CONTABLE2 <> ? AND TAB_NUMTAB = ? ORDER BY TAB_NUMTAB "
Set PSMA_BUSCA = CN.CreateQuery("", pub_cadena)
PSMA_BUSCA(0) = 0
PSMA_BUSCA(1) = 0
PSMA_BUSCA(2) = 0
PSMA_BUSCA(3) = 0
Set man_busca = PSMA_BUSCA.OpenResultset(rdOpenKeyset, rdConcurValues)


TURNOS.Cols = 4
TURNOS.Rows = 1
TURNOS.ColWidth(0) = 500
TURNOS.ColWidth(1) = 1900
TURNOS.ColWidth(2) = 0
TURNOS.ColWidth(3) = 0


PUB_TIPREG = 2102
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_TAB_LLAVE
fila = 0
Do Until tab_mayor.EOF
     TURNOS.Rows = TURNOS.Rows + 1
     TURNOS.RowHeight(fila) = 285
     TURNOS.TextMatrix(fila, 0) = Format(tab_mayor!tab_numtab, "000")
     TURNOS.TextMatrix(fila, 1) = Trim(tab_mayor!tab_nomlargo)
     TURNOS.TextMatrix(fila, 2) = Trim(tab_mayor!tab_codART)
     fila = fila + 1
     tab_mayor.MoveNext
Loop
TURNOS.Rows = TURNOS.Rows - 1
loc_key = 0
LIMPIA_VEN
BLOQUEA_TEXT txtnombre, serie_g, numfac_g, Serie_b, numfac_b, serie_f, numfac_f, numfac_p, numfac_p_f, serie_p
BLOQUEA_TEXT numfac_g_f, numfac_b_f, numfac_f_f, cheguia, cheboleta, chefactura, Check1, productos, lisresp, TURNOS
txt_key.Enabled = True
remi.Enabled = False
If LK_FLAG_FACTURACION = "V" Then
  F2.Visible = True
End If
flag_tempo = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
ws_conta = 0
End Sub

Public Sub BLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = False
' o1.BackColor = QBColor(7)
End If
If Not IsMissing(o2) Then
 o2.Enabled = False
 'o2.BackColor = QBColor(7)
End If
If Not IsMissing(o3) Then
 o3.Enabled = False
 'o3.BackColor = QBColor(7)
End If
If Not IsMissing(o4) Then
 o4.Enabled = False
 'o4.BackColor = QBColor(7)
End If
If Not IsMissing(o5) Then
 o5.Enabled = False
 'o5.BackColor = QBColor(7)
End If
If Not IsMissing(o6) Then
 o6.Enabled = False
 'o6.BackColor = QBColor(7)
End If
If Not IsMissing(o7) Then
 o7.Enabled = False
 'o7.BackColor = QBColor(7)
End If
If Not IsMissing(o8) Then
 o8.Enabled = False
 'o8.BackColor = QBColor(7)
End If
If Not IsMissing(o9) Then
 o9.Enabled = False
 'o9.BackColor = QBColor(7)
End If
If Not IsMissing(o10) Then
 o10.Enabled = False
 'o10.BackColor = QBColor(7)
End If
End Sub
Public Sub DESBLOQUEA_TEXT(Optional o1, Optional o2, Optional o3, Optional o4, Optional o5, Optional o6, Optional o7, Optional o8, Optional o9, Optional o10)
'** BLOQUEA TEXTBOX  CANTIDAD DE OBJECTOS **
If Not IsMissing(o1) Then
 o1.Enabled = True
' o1.BackColor = QBColor(15)
End If
If Not IsMissing(o2) Then
 o2.Enabled = True
' o2.BackColor = QBColor(15)
End If
If Not IsMissing(o3) Then
 o3.Enabled = True
' o3.BackColor = QBColor(15)
End If
If Not IsMissing(o4) Then
 o4.Enabled = True
' o4.BackColor = QBColor(15)
End If
If Not IsMissing(o5) Then
 o5.Enabled = True
' o5.BackColor = QBColor(15)
End If
If Not IsMissing(o6) Then
 o6.Enabled = True
' o6.BackColor = QBColor(15)
End If
If Not IsMissing(o7) Then
 o7.Enabled = True
' o7.BackColor = QBColor(15)
End If
If Not IsMissing(o8) Then
 o8.Enabled = True
' o8.BackColor = QBColor(15)
End If
If Not IsMissing(o9) Then
 o9.Enabled = True
' o9.BackColor = QBColor(15)
End If
If Not IsMissing(o10) Then
 o10.Enabled = True
' o10.BackColor = QBColor(15)
End If
End Sub

Private Sub gridpro_EnterCell()
textovar.Visible = False
textovar.Text = Trim(gridpro.TextMatrix(gridpro.Row, gridpro.COL))
textovar.Left = gridpro.Left + gridpro.CellLeft
textovar.Width = gridpro.CellWidth
textovar.Height = gridpro.CellHeight
textovar.Top = gridpro.Top + gridpro.CellTop  ' ESTADO.Top + gridpro.Top + gridpro.CellTop - 1560 '480
If gridpro.COL = 1 Then
 If Trim(gridpro.TextMatrix(gridpro.Row, 1)) = "" Then
   If Trim(gridpro.TextMatrix(gridpro.Row, 0)) <> "" Then
    textovar.Visible = True
    textovar.SetFocus
   End If
 End If
End If
If gridpro.COL = 0 Then
 If Trim(gridpro.TextMatrix(gridpro.Row, 0)) = "" Then
  textovar.Visible = True
  textovar.SetFocus
 End If
End If


End Sub



Private Sub gridpro_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, wc
Dim wprecios As String * 12
Static CONS
Dim wactivo As Integer
If KeyAscii <> 13 Then Exit Sub
If gridpro.Rows <= 1 Then Exit Sub
If gridpro.COL >= 2 Then
   If gridpro.COL <> 5 Then
    Exit Sub
   End If
End If

    textovar.Visible = True
    Azul textovar, textovar
    textovar.SetFocus
Exit Sub
pasa:
Resume Next

End Sub

Private Sub gridpro_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wc
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean

If KeyCode = 45 Then
   If Trim(gridpro.TextMatrix(gridpro.Rows - 1, 2)) = "" Then Exit Sub
   gridpro.Rows = gridpro.Rows + 1
   gridpro.COL = 0
   gridpro.RowHeight(gridpro.Rows - 1) = 285
   gridpro.COL = 1
   gridpro.COL = 0
   gridpro.Row = gridpro.Rows - 1
   Exit Sub
End If

If KeyCode = 46 Then
If gridpro.Rows <= 2 Then Exit Sub
If gridpro.Rows <= 3 Then
    pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
    If pub_mensaje = vbNo Then
      gridpro.SetFocus
      Exit Sub
    End If
Else
   pub_mensaje = MsgBox("Quitar el Surtidor ? ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridpro.SetFocus
     Exit Sub
   Else
   '  gridpro.RowHeight(gridpro.Row) = 1
   gridpro.RemoveItem (gridpro.Row)
   gridpro.Row = gridpro.Row
   gridpro.Refresh
   gridpro.SetFocus
   End If
End If
End If
'gridpro.SetFocus
Exit Sub



End Sub

Private Sub lisresp_Click()
 If flag_tempo = "A" Then
 Else
    TURNOS.TextMatrix(TURNOS.Row, 2) = Trim(Right(lisresp.Text, 8))
 End If
 
End Sub

Private Sub lisresp_KeyPress(KeyAscii As Integer)
If TURNOS.Enabled Then TURNOS.SetFocus
End Sub

Private Sub lisresp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
 PUB_TIPREG = 2103
 PUB_CODCIA = LK_CODCIA
 Load FrmDatplac
 FrmDatplac.Caption = "Responsables"
 FrmDatplac.Show 1
 lisresp.Clear
 PUB_TIPREG = 2103
 PUB_CODCIA = LK_CODCIA
 SQ_OPER = 2
 LEER_TAB_LLAVE
 fila = 0
 Do Until tab_mayor.EOF
   lisresp.AddItem tab_mayor!tab_nomlargo + String(80, " ") + Format(tab_mayor!tab_numtab, "000")
   tab_mayor.MoveNext
 Loop
TURNOS.SetFocus
TURNOS_EnterCell
 'DoEvents
 'PSVE2_LLAVE(0) = LK_CODCIA
 'VE2_LLAVE.Requery
 'lisresp.Clear
 'Do Until VE2_LLAVE.EOF
 '  lisresp.AddItem VE2_LLAVE!VE2_NOMBRE + String(60, " ") + Str(VE2_LLAVE!VE2_CODVEN)
 '  VE2_LLAVE.MoveNext
 'Loop
 'lisresp.SetFocus
End If
End Sub

Private Sub ListView1_GotFocus()
If loc_key <> 0 Then
 Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
 ListView1.ListItems.Item(loc_key).Selected = True
 ListView1.ListItems.Item(loc_key).EnsureVisible
End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView1.SelectedItem.Index
 txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 txt_key.Text = ""
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
txt_key_KeyPress 13
End Sub

Private Sub ListView2_GotFocus()
If loc_key <> 0 Then
 Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
 ListView2.ListItems.Item(loc_key).Selected = True
 ListView2.ListItems.Item(loc_key).EnsureVisible
End If


End Sub

Private Sub ListView2_ItemClick(ByVal Item As ComctlLib.ListItem)
If loc_key <> 0 Then
 loc_key = ListView2.SelectedItem.Index
 txt_key.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
End If

End Sub

Private Sub ListView2_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 txt_arti.Text = ""
End If
If KeyAscii <> 13 Then
 Exit Sub
End If
txt_arti_KeyPress 13

End Sub

Private Sub numfac_b_f_GotFocus()
Azul numfac_b_f, numfac_b_f
End Sub

Private Sub numfac_b_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 serie_f.SetFocus
End If

End Sub

Private Sub numfac_b_GotFocus()
Azul numfac_b, numfac_b
End Sub

Private Sub numfac_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b_f.SetFocus
End If

End Sub

Private Sub numfac_f_f_GotFocus()
Azul numfac_f_f, numfac_f_f
End Sub

Private Sub numfac_f_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then
  Exit Sub
End If
If cmdmodificar.Enabled Then
   cmdmodificar.SetFocus
Else
   CmdAgregar.SetFocus
End If

End Sub

Private Sub numfac_f_GotFocus()
Azul numfac_f, numfac_f
End Sub

Private Sub numfac_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  numfac_f_f.SetFocus
End If
End Sub

Private Sub numfac_g_f_GotFocus()
Azul numfac_g_f, numfac_g_f
End Sub

Private Sub numfac_g_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 Serie_b.SetFocus
End If

End Sub

Private Sub numfac_g_GotFocus()
Azul numfac_g, numfac_g
End Sub

Private Sub numfac_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g_f.SetFocus
End If

End Sub

Private Sub numfac_p_f_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul serie_g, serie_g
End If
End Sub

Private Sub numfac_p_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul numfac_p_f, numfac_p_f
End If
End Sub

Private Sub productos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  TURNOS.SetFocus
  Exit Sub
End If

End Sub

Private Sub productos_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then
   frapro.Visible = True
   txt_arti.Text = ""
   txt_arti.SetFocus
End If
If KeyCode = 46 And productos.ListIndex > -1 Then
    productos.RemoveItem productos.ListIndex
    productos.SetFocus
End If

End Sub

Private Sub remi_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
End Sub

Private Sub Serie_b_GotFocus()
Azul Serie_b, Serie_b
End Sub

Private Sub Serie_b_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_b.SetFocus
End If

End Sub

Private Sub serie_f_GotFocus()
Azul serie_f, serie_f
End Sub

Private Sub serie_f_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_f.SetFocus
End If

End Sub

Private Sub serie_g_GotFocus()
'Azul serie_g, serie_g
End Sub

Private Sub serie_g_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 numfac_g.SetFocus
End If
End Sub

Private Sub serie_p_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Azul numfac_p, numfac_p
End If
End Sub

Private Sub textvar_Change()

End Sub

Private Sub textvar_KeyPress(KeyAscii As Integer)
End Sub

Private Sub TURNOS_EnterCell()
Dim WCODI
If TURNOS.COL <> 1 Then
  TURNOS.COL = 1
End If
TURNOS.CellBackColor = QBColor(11)
If lisresp.ListCount = 0 Then
  MsgBox "NO Existen nigun Responsables. Ingresar...", 48, Pub_Titulo
  If lisresp.Enabled And lisresp.Visible Then lisresp.SetFocus
  Exit Sub
End If
WCODI = TURNOS.TextMatrix(TURNOS.Row, 2)
lisresp.Visible = False
flag_tempo = "A"
For fila = 0 To lisresp.ListCount - 1
   lisresp.ListIndex = fila
   If Val(WCODI) = Val(Trim(Right(lisresp.Text, 8))) Then
     lisresp.Visible = True
     flag_tempo = ""
     Exit Sub
   End If
Next fila
flag_tempo = ""
pasa:
lisresp.ListIndex = -1
lisresp.Visible = True


End Sub

Private Sub TURNOS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If serie_p.Visible Then serie_p.SetFocus
End If

End Sub

Private Sub TURNOS_LeaveCell()
If TURNOS.COL = 1 Then TURNOS.CellBackColor = QBColor(15)
End Sub

Private Sub txt_arti_LostFocus()
 'frapro.Visible = False
 'productos.SetFocus
End Sub

Private Sub txt_key_GotFocus()
 Azul txt_key, txt_key
End Sub
Private Sub txt_key_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_key.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo FIN
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
GoTo FIN
POSICION:
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txt_key.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_key.SelStart = Len(txt_key.Text)
  DoEvents
FIN:

End Sub
Private Sub txt_key_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
If KeyAscii = 27 And Trim(txtnombre.Text) = "" Then
 txt_key.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo FIN
End If
pu_codclie = Val(txt_key.Text)
If Len(txt_key.Text) = 0 Or txt_key.Locked Then
   Exit Sub
End If
If pu_codclie <> 0 And IsNumeric(txt_key.Text) = True Then
   loc_key = 0
   On Error GoTo mucho
   PUB_CODVEN = Val(txt_key.Text)
   pu_codcia = LK_CODCIA
   SQ_OPER = 1
   LEER_VEN_LLAVE
   On Error GoTo 0
   If ven_llave.EOF Then
     Azul txt_key, txt_key
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txt_key.SetFocus
     GoTo FIN
   End If
   ListView1.Visible = False
   cmdcancelar.Enabled = True
   LLENA_VEN 0
   txt_key.Locked = True
   cmdmodificar.SetFocus
   Screen.MousePointer = 0
   Exit Sub
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_key.Text)) = Left(valor, Len(Trim(txt_key.Text))) Then
   Else
      Exit Sub
   End If
   ListView1.Visible = False
   cmdcancelar.Enabled = True
   LLENA_VEN 0
    txt_key.Locked = True
   cmdcancelar.Enabled = True
    cmdmodificar.SetFocus
End If
dale:
mucho:
ListView1.Visible = False
FIN:
End Sub

Private Sub txt_key_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
If Len(txt_key.Text) = 0 Or txt_key.Locked = True Or IsNumeric(txt_key.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txt_key.Text) = 1 Then
    VAR = Asc(txt_key.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 0
    archi = "SELECT * FROM VEMAEST WHERE  VEM_CODCIA = '" & LK_CODCIA & "' AND VEM_NOMBRE BETWEEN '" & txt_key.Text & "' AND  '" & VAR & "' ORDER BY VEM_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If ListView1.Visible = False Then
        loc_key = 0
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txt_key.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If
End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True 'Not LblMensaje.Visible
 If CU > 8 Then
   CU = 0
   PARPADEA.Enabled = False
   LblMensaje.Visible = False
 End If
End Sub


Private Sub txtnombre_GotFocus()
Azul txtnombre, txtnombre
End Sub

Private Sub txtnombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 gridpro.COL = 0
 If gridpro.Cols = 1 Then
   gridpro.COL = 1
 End If
 gridpro.SetFocus
End If
End Sub

Public Function ES_FECHAS(CAMPOFECHA As MaskEdBox) As String
Dim wfecha As String
ES_FECHAS = "0"
If CAMPOFECHA = "00/00/0000" Then
 Exit Function
End If
If Right(CAMPOFECHA.Text, 2) = "__" Then
  wfecha = Left(CAMPOFECHA.Text, 8)
Else
  wfecha = Trim(CAMPOFECHA.Text)
End If
If Not IsDate(wfecha) Then
  ES_FECHAS = "1"
  Exit Function
End If
ES_FECHAS = wfecha
End Function

Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Integer)
For fila = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(fila), 3))) = txt Then
        WCONTROL.ListIndex = fila
        Exit Sub
    End If
Next fila

End Sub


' inicio
Private Sub txt_arti_Change()
If txt_arti.Text = "" Then
 VAR_ACTIVAR = 0
End If
End Sub

Private Sub txt_arti_GotFocus()
'PRO_GOTFOCUS
End Sub

Private Sub txt_arti_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_arti.Text = "" Then
  loc_key = 1
  Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  GoTo FIN
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo FIN
POSICION:
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  txt_arti.Text = Trim(ListView2.ListItems.Item(loc_key).Text)
  DoEvents
  txt_arti.SelStart = Len(txt_arti.Text)
  DoEvents
FIN:

End Sub
Private Sub txt_arti_KeyPress(KeyAscii As Integer)
Dim wwarti As String * 8
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 If txt_arti.Text <> "" Then
   ListView2.Visible = False
   txt_arti.Text = ""
  Else
    productos.SetFocus
    frapro.Visible = False
    ListView2.Visible = False
  End If
End If
If KeyAscii <> 13 Then
   GoTo FIN
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 PUB_KEY = Val(txt_arti.Text)
 If Len(txt_arti.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(txt_arti.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = txt_arti.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul txt_arti, txt_arti
       GoTo FIN
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    productos.AddItem art_LLAVE!art_nombre + String(80, " ") + Str(art_LLAVE!ART_KEY)
    productos.SetFocus
    frapro.Visible = False
    ListView2.Visible = False
    
    Exit Sub
Else
  If ListView2.Visible = False And VAR_ACTIVAR <> 99 And txt_arti.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = txt_arti.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul txt_arti, txt_arti
       Exit Sub
     End If
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     productos.AddItem art_llave_alt!art_nombre + String(80, " ") + Str(art_llave_alt!ART_KEY)
      productos.SetFocus
     frapro.Visible = False
     ListView2.Visible = False
     Exit Sub
  Else
    If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView2.ListItems.Item(loc_key).Text)
    If Trim(UCase(txt_arti.Text)) = Left(valor, Len(Trim(txt_arti.Text))) And Len(Trim(txt_arti.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        txt_arti.Text = Trim(ListView2.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       txt_arti.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
      Else
       txt_arti.Text = Trim(ListView2.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView2.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = txt_arti.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul txt_arti, txt_arti
        Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      productos.AddItem art_LLAVE!art_nombre + String(80, " ") + Str(art_LLAVE!ART_KEY)
      productos.SetFocus
      frapro.Visible = False
      ListView2.Visible = False
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView2.Visible = False
FIN:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_arti, txt_arti
End Sub

Private Sub txt_arti_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
Dim ws_codcia As String * 2
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(txt_arti.Text) = 0 Or txt_arti.Text = "" Then
    ListView2.Visible = False
    Exit Sub
  End If
  If txt_arti.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf txt_arti.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(txt_arti.Text, 1) = "*" Then
   txt_arti.Text = Mid(txt_arti.Text, 2, Len(txt_arti.Text))
   txt_arti.SelStart = Len(txt_arti.Text)
  End If
Else
 If Len(txt_arti.Text) = 0 Or IsNumeric(txt_arti.Text) = True Then
   ListView2.Visible = False
   Exit Sub
 End If
End If
If ListView2.Visible = False And KeyCode <> 13 Or Len(txt_arti.Text) = 1 Then
    VAR = Asc(txt_arti.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    ws_codcia = LK_CODCIA
    If LK_EMP_PTO = "A" Then
      ws_codcia = "00"
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_KEY <> 0 AND ART_CALIDAD = " & 1 & " AND ART_CODCIA = '" & ws_codcia & "' AND ART_ALTERNO BETWEEN '" & txt_arti.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE ART_KEY <> 0 AND ART_CALIDAD = " & 1 & " AND  ART_CODCIA = '" & ws_codcia & "' AND ART_NOMBRE BETWEEN '" & txt_arti.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
     PROC_LISVIEW ListView2
     loc_key = 0
     If ListView2.Visible Then
      loc_key = 1
     End If
     Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(txt_arti.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView2.ListItems.count Then
      ListView2.ListItems.Item(ListView2.ListItems.count).EnsureVisible
   Else
     ListView2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If

End Sub
Public Sub cabe_pro()
    gridpro.Clear
    gridpro.Cols = 6
    gridpro.ColWidth(0) = 700
    gridpro.ColWidth(1) = 1000
    gridpro.ColWidth(2) = 2500
    gridpro.ColWidth(3) = 0
    gridpro.ColWidth(4) = 0
    gridpro.ColWidth(5) = 800
    gridpro.TextMatrix(0, 0) = "Surtidor"
    gridpro.TextMatrix(0, 1) = "Codigo"
    gridpro.TextMatrix(0, 2) = "Descripción Producto"
    gridpro.TextMatrix(0, 5) = "Lectura"
    
    gridpro.Rows = 1
End Sub

' comienza texto var

Private Sub textovar_Change()
If Not textovar.Visible Then Exit Sub
If gridpro.COL = 1 Then
    gridpro.TextMatrix(gridpro.Row, 1) = ""
    gridpro.TextMatrix(gridpro.Row, 2) = ""
    gridpro.TextMatrix(gridpro.Row, 3) = ""
Else
 If gridpro.COL = 0 Then
   gridpro.Text = Format(textovar.Text, "0")
 ElseIf gridpro.COL = 1 Then
   gridpro.Text = textovar.Text
 Else
   gridpro.Text = Format(textovar.Text, "0.00")
 End If
End If
End Sub
Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)
' busca arti
If Not ListView2.Visible Then
If KeyCode = 40 Then  ' flecha abajo
  If gridpro.Row = gridpro.Rows - 1 Then Exit Sub
  If Trim(gridpro.Text) <> "" Then Exit Sub
  gridpro.Row = gridpro.Row + 1
  gridpro.SetFocus
  Exit Sub
End If
If KeyCode = 38 Then
 If Trim(gridpro.Text) <> "" Then Exit Sub
 gridpro.Row = gridpro.Row - 1
 gridpro.SetFocus
 Exit Sub
End If
If KeyCode = 39 Then
If Trim(gridpro.Text) <> "" Then Exit Sub
 gridpro.COL = gridpro.COL + 1
 gridpro.SetFocus
 Exit Sub
End If
End If
If gridpro.COL <> 1 Then Exit Sub
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And textovar.Text = "" Then
  loc_key = 1
  Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
'  ListView2.Visible = False
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  GoTo FIN
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo FIN
POSICION:
'  KeyCode = 0
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  textovar.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  textovar.SelStart = Len(textovar.Text)
  DoEvents
FIN:

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
'SOLO_DECIMAL TEXTOVAR, KeyAscii
If KeyAscii = 27 Then
  If textovar.Text = "" Then
    textovar.Visible = False
    gridpro.SetFocus
    Exit Sub
  End If
  textovar.Visible = False
  ListView2.Visible = False
  Exit Sub
End If
'If gridpro.COL = 2 Or gridpro.COL = 4 Or gridpro.COL = 5 Then Consistencias gridpro, textovar, KeyAscii
If gridpro.COL = 0 Then
  SOLO_ENTERO KeyAscii
End If
If gridpro.COL = 5 Then
  SOLO_DECIMAL textovar, KeyAscii
End If

If KeyAscii <> 13 Then Exit Sub

If gridpro.COL = 0 Then
 textovar.Visible = False
 If gridpro.Row < gridpro.Rows - 1 Then
     gridpro.Row = gridpro.Row + 1
 End If
 If Trim(gridpro.Text) <> "" Then
   gridpro.SetFocus
   Exit Sub
 End If
 Exit Sub
End If
If gridpro.COL = 0 Then
 textovar.Visible = False
 If Trim(gridpro.TextMatrix(gridpro.Rows - 1, 1)) <> "" Then
   gridpro.Rows = gridpro.Rows + 1
   gridpro.RowHeight(gridpro.Rows - 1) = 285
   gridpro.Row = gridpro.Rows - 1
 Else
  If gridpro.Row < gridpro.Rows - 1 Then
     gridpro.Row = gridpro.Row + 1
  End If
 End If
 gridpro.COL = 1
 textovar.Visible = True
 textovar.SetFocus
 Exit Sub
End If




If gridpro.COL <> 1 Then
  textovar.Visible = False
  Exit Sub
End If

Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView2.Visible = False
 textovar.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo FIN
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(textovar.Text)
 On Error GoTo 0
 If Len(textovar.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(textovar.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = textovar.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul3 textovar, textovar
       GoTo FIN
    End If
    If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 textovar, textovar
       GoTo FIN
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = WCOD_ORIGINAL
    LEER_ARM_LLAVE
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_SECUEN = 0
    LEER_PRE_LLAVE
    gridpro.TextMatrix(gridpro.Row, 16) = pre_llave!pre_UNIDAD
    gridpro.TextMatrix(gridpro.Row, 15) = arm_llave!ART_stock
    gridpro.TextMatrix(gridpro.Row, 12) = pre_llave!pre_equiv
    gridpro.TextMatrix(gridpro.Row, 11) = pre_llave!pre_secuencia
    gridpro.TextMatrix(gridpro.Row, 0) = art_LLAVE!art_nombre
    gridpro.TextMatrix(gridpro.Row, 10) = art_LLAVE!ART_KEY
    ListView2.Visible = False
    textovar.Visible = False
    gridpro.COL = 2
    If Trim(gridpro.Text) <> "" Then
      gridpro.SetFocus
      Exit Sub
    End If
    textovar.Visible = True
    textovar.SetFocus
    Exit Sub
Else
  If ListView2.Visible = False And VAR_ACTIVAR <> 99 And textovar.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = textovar.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul textovar, textovar
       Exit Sub
     End If
     If art_llave_alt!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul textovar, textovar
       GoTo FIN
     End If
     ListView2.Visible = False
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_KEY = WCOD_ORIGINAL
     LEER_ART_LLAVE
     gridpro.TextMatrix(gridpro.Row, 1) = art_LLAVE!art_alterno
     gridpro.TextMatrix(gridpro.Row, 2) = art_LLAVE!art_nombre
     gridpro.TextMatrix(gridpro.Row, 3) = art_LLAVE!ART_KEY
     textovar.Visible = False
     ListView2.Visible = False
     gridpro.COL = 1
     If Trim(gridpro.Text) <> "" Then
       gridpro.SetFocus
       Exit Sub
     End If
     textovar.Visible = True
     Azul3 textovar, textovar
     Exit Sub
  Else
    If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView2.ListItems.Item(loc_key).Text)
    If Trim(UCase(textovar.Text)) = Left(valor, Len(Trim(textovar.Text))) And Len(Trim(textovar.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        textovar.Text = Trim(ListView2.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       textovar.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
      Else
       textovar.Text = Trim(ListView2.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView2.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = textovar.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul3 textovar, textovar
        Exit Sub
      End If
      If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 textovar, textovar
       GoTo FIN
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_CODART = WCOD_ORIGINAL
      LEER_ARM_LLAVE
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_SECUEN = 0
      LEER_PRE_LLAVE
      gridpro.TextMatrix(gridpro.Row, 16) = pre_llave!pre_UNIDAD
      gridpro.TextMatrix(gridpro.Row, 15) = arm_llave!ART_stock
      gridpro.TextMatrix(gridpro.Row, 12) = pre_llave!pre_equiv
      gridpro.TextMatrix(gridpro.Row, 11) = pre_llave!pre_secuencia
      ListView2.Visible = False
      gridpro.TextMatrix(gridpro.Row, 0) = art_LLAVE!art_nombre
      gridpro.TextMatrix(gridpro.Row, 10) = art_LLAVE!ART_KEY
      gridpro.COL = 2
      If Trim(gridpro.Text) <> "" Then
        gridpro.SetFocus
        Exit Sub
      End If
      textovar.Visible = True
      textovar.SetFocus
     
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView2.Visible = False
FIN:
mucho:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul3 textovar, textovar
  

Exit Sub

End Sub

Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)
If gridpro.COL <> 1 Then Exit Sub
' busca arti
Dim VAR
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(textovar.Text) = 0 Or Trim(textovar.Text) = "" Then
    ListView2.Visible = False
    Exit Sub
  End If
  If textovar.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf textovar.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(textovar.Text, 1) = "*" Then
   textovar.Text = Mid(textovar.Text, 2, Len(textovar.Text))
   textovar.SelStart = Len(textovar.Text)
  End If
Else
 If Len(textovar.Text) = 0 Or IsNumeric(textovar.Text) = True Then
   ListView2.Visible = False
   Exit Sub
 End If
End If
If ListView2.Visible = False And KeyCode <> 13 Or Len(textovar.Text) = 1 Then
    VAR = Asc(textovar.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & textovar.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
   ' If Len(TEXTOVAR.text) > 1 And ListView2.ListItems.count = 0 Then
   ' Else
     PROC_LISVIEW ListView2
   ' End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(textovar.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView2.ListItems.count Then
      ListView2.ListItems.Item(ListView2.ListItems.count).EnsureVisible
   Else
     ListView2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub



Public Function CONS_SUR() As String
Dim Wflag As String * 1
Dim WMEN As String

PSMA_BUSCA(0) = LK_CODCIA
PSMA_BUSCA(1) = 2105
PSMA_BUSCA(2) = Val(txt_key.Text)
WMEN = "Surtidores ya Existen...!!!" & Chr(13)
Wflag = ""
For fila = 1 To gridpro.Rows - 1
 If Val(gridpro.TextMatrix(fila, 0)) = 0 Then
  WMEN = "Falta ingregar datos, verificar"
  Wflag = "A"
  GoTo SALE
 End If
 PSMA_BUSCA(3) = gridpro.TextMatrix(fila, 0)
 man_busca.Requery
 If Not man_busca.EOF Then
   WMEN = WMEN + "Sut.: " + Trim(gridpro.TextMatrix(fila, 0)) + " pertenece a Isla: " + man_busca!TAB_contable2 + Chr(13)
   Wflag = "A"
 End If
Next fila
SALE:
CONS_SUR = ""
If Wflag = "A" Then
  CONS_SUR = WMEN
End If

End Function
