VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmComaest 
   Caption         =   "Cuentas ..."
   ClientHeight    =   5628
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8472
   Icon            =   "frmcomaest.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5628
   ScaleWidth      =   8472
   WindowState     =   2  'Maximized
   Begin VB.Frame FLISTA 
      Caption         =   "Cuentas   ... "
      Height          =   3615
      Left            =   4800
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox LEIDO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   2775
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   3975
         _ExtentX        =   7006
         _ExtentY        =   4890
         _Version        =   327680
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   3000
      TabIndex        =   40
      Top             =   1920
      Width           =   2055
      Begin VB.TextBox Txtcom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   42
         Text            =   "com_cuenta_autom_d"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Txtcom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   480
         TabIndex        =   41
         Text            =   "com_cuenta_auto_h"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label ACT_PAS 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   1680
         Width           =   75
      End
      Begin VB.Label SALDOACTUAL 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   2040
         Width           =   75
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Aut. Debe"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Aut. Haber"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame FMODO 
      Caption         =   "Modo de :"
      Height          =   855
      Left            =   7080
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label Modo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2055
      End
   End
   Begin Crystal.CrystalReport ListaRepo 
      Left            =   240
      Top             =   6240
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   262150
   End
   Begin VB.CommandButton CmdRepo 
      Caption         =   "Re&portes Cuentas"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   1080
      TabIndex        =   33
      Top             =   6360
      Visible         =   0   'False
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid ListExiste 
      Height          =   735
      Left            =   120
      TabIndex        =   23
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   656
      _ExtentY        =   1291
      _Version        =   327680
   End
   Begin VB.Frame F3 
      Height          =   2415
      Left            =   5280
      TabIndex        =   21
      Top             =   1920
      Width           =   3375
      Begin VB.TextBox Txtcom 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   1680
         TabIndex        =   37
         Text            =   "com_deb_mes"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Txtcom 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   1680
         TabIndex        =   36
         Text            =   "com_hab_mes"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Txtcom 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Text            =   "com_cuenta_sup"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Txtcom 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   8
         Text            =   "com_sal_inicial"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Txtcom 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Text            =   "com_flag_afectacion"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Debe Mes"
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   39
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Haber Mes"
         Height          =   255
         Index           =   10
         Left            =   1680
         TabIndex        =   38
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cta. Superior"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Inicial"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Flag.. Afectación"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Timer PARPADEA 
      Interval        =   500
      Left            =   7800
      Top             =   3480
   End
   Begin VB.CommandButton CmdCerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame F2 
      Height          =   2415
      Left            =   0
      TabIndex        =   18
      Top             =   1920
      Width           =   2895
      Begin VB.ComboBox com_tipo_cuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox com_signo_d 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox com_signo_h 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Cuenta"
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Signo Debe"
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Signo Haber"
         Height          =   255
         Index           =   13
         Left            =   1440
         TabIndex        =   27
         Top             =   1080
         Width           =   975
      End
   End
   Begin VB.Frame F1 
      Caption         =   "Cuenta .."
      Height          =   1695
      Left            =   0
      TabIndex        =   15
      Top             =   120
      Width           =   6975
      Begin VB.TextBox TxtCtaSobre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         MaxLength       =   1
         TabIndex        =   47
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox Txtcom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   5640
         TabIndex        =   2
         Text            =   "com_nivel"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox Txtcom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   240
         MaxLength       =   40
         TabIndex        =   3
         Text            =   "com_descripcion"
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox Txtcom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   0
         Text            =   "com_cuenta"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Txtcom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3360
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "com_codcia"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cta. Sobregiro"
         Height          =   195
         Left            =   5640
         TabIndex        =   48
         Top             =   960
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel :"
         Height          =   255
         Index           =   14
         Left            =   4920
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Cia.:"
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Frame REPLACE1 
      Caption         =   "Reemplasar los Signos"
      Height          =   615
      Left            =   6840
      TabIndex        =   49
      Top             =   4320
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton Command1 
         Caption         =   "Reemplasar"
         Height          =   315
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   1815
      End
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   3720
      TabIndex        =   51
      Top             =   4560
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5101
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   600
      X2              =   8640
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   600
      X2              =   8640
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label LblMensaje 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   20
      Top             =   4800
      Width           =   420
   End
End
Attribute VB_Name = "FrmComaest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LOC_CTA_SUP As String
Dim LOC_TIPO_CTA As Integer
Dim LOC_SIG_DEB As Integer
Dim LOC_SIG_HAB As Integer
Dim CU As Integer
Dim pasa As Boolean
Dim wCOM_NIVEL(6) As Integer
Dim NIVEL_ACT As Integer
Dim NIVEL_MAX As Integer
Dim IND As Integer
Dim Wflag As String * 1
Dim llave1
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery


Public Sub GRABAR_COM()
Dim i As Integer
Dim wcta As Integer
If Left(cmdModificar.Caption, 2) = "&G" Then
   com_llave.Edit
Else
   com_llave.AddNew
   com_llave!COM_CODCIA = Trim(FrmComaest.Txtcom(0).Text)
   com_llave!com_cuenta = Trim(FrmComaest.Txtcom(1).Text)
   com_llave!COM_NIVEL = NIVEL_ACT
End If

com_llave.rdoColumns(2) = Trim(FrmComaest.Txtcom(2).Text)
com_llave.rdoColumns(3) = Trim(FrmComaest.Txtcom(3).Text)
com_llave.rdoColumns(4) = Trim(FrmComaest.Txtcom(4).Text)
com_llave.rdoColumns(6) = Trim(FrmComaest.Txtcom(6).Text)
com_llave.rdoColumns(7) = Trim(FrmComaest.Txtcom(7).Text)
com_llave.rdoColumns(8) = " "
com_llave!COM_TIPO_CTA = Val(Right(FrmComaest.com_tipo_cuenta.Text, 3))
com_llave!com_SAL_INICIAL = Val(FrmComaest.Txtcom(10).Text)
com_llave!COM_DEB_MES = Val(FrmComaest.Txtcom(11).Text)
com_llave!COM_HAB_MES = Val(FrmComaest.Txtcom(12).Text)
com_llave!com_signo_d = Val(FrmComaest.com_signo_d)
com_llave!com_signo_h = Val(FrmComaest.com_signo_h)
com_llave!com_flag_sobregiro = TxtCtaSobre.Text
wcta = Val(Right(FrmComaest.com_tipo_cuenta.Text, 3))
Select Case wcta
Case 1, 3
 com_llave!COM_act_pas = 1
Case 2, 4
 com_llave!COM_act_pas = 2
Case 5
 com_llave!COM_act_pas = 3
Case Else
 com_llave!COM_act_pas = 0
End Select
com_llave.Update
End Sub
Public Sub MENSAJE_COM(TEXTO As String)
  LblMensaje.Caption = TEXTO
  Parpadea.Enabled = True
End Sub

Public Sub LLENA_COM(ban As Integer)
Dim i As Integer
Dim wactual As Currency
If ban = 0 Then
       grid1.COL = 0
       PUB_CUENTA = Trim(FrmComaest.grid1.Text)
       SQ_OPER = 1
       LEER_COM_LLAVE
End If

For i = 0 To 12
  If i <> 9 And i <> 8 Then
   If i <> 0 And i <> 1 Then
    FrmComaest.Txtcom(i).Text = Nulo_Valors(Trim(com_llave.rdoColumns(i)))
   End If
  End If
Next i
FrmComaest.Txtcom(0).Text = Nulo_Valors(Trim(com_llave!COM_CODCIA))
FrmComaest.Txtcom(1).Text = Nulo_Valors(Trim(com_llave!com_cuenta))
ASIGNA_SIGNO com_signo_d, Nulo_Valor0(com_llave!com_signo_d)
ASIGNA_SIGNO com_signo_h, Nulo_Valor0(com_llave!com_signo_h)
ASIGNA_INT com_tipo_cuenta, Nulo_Valor0(com_llave!COM_TIPO_CTA)
TxtCtaSobre.Text = Nulo_Valors(Trim(com_llave!com_flag_sobregiro))
ACT_PAS.Caption = "Act_Pas.  = " & CStr(Nulo_Valor0(com_llave!COM_act_pas))
wactual = Nulo_Valor0(com_llave!com_SAL_INICIAL) + Nulo_Valor0(com_llave!COM_DEB_MES) * Nulo_Valor0(com_llave!com_signo_d) + Nulo_Valor0(com_llave!COM_HAB_MES) * Nulo_Valor0(com_llave!com_signo_h)
SALDOACTUAL.Caption = "Saldo Act.= " & Format(wactual, "##,###,##0.00")
End Sub
Public Sub ASIGNA_INT(WCONTROL As ComboBox, txt As Integer)
Dim C As Integer
For C = 0 To WCONTROL.ListCount - 1
    If Val(Trim(Right(WCONTROL.List(C), 3))) = txt Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub

Public Sub ASIGNA_SIGNO(WCONTROL As ComboBox, txt As Integer)
Dim C As Integer
For C = 0 To WCONTROL.ListCount - 1
    If Val(WCONTROL.List(C)) = txt Then
        WCONTROL.ListIndex = C
        Exit Sub
    End If
Next C
End Sub

Public Sub LIMPIA_COM()
Dim i As Integer
For i = 0 To 12
  If i <> 9 And i <> 8 Then
      Txtcom(i).Text = ""
  End If
Next i
TxtCtaSobre.Text = ""
SALDOACTUAL.Caption = ""
ACT_PAS.Caption = ""
com_signo_d.ListIndex = -1
com_signo_h.ListIndex = -1
com_tipo_cuenta.ListIndex = -1

End Sub

Private Sub cmdagregar_Click()
On Error GoTo SALE
If Left(cmdAgregar.Caption, 2) = "&A" Then
    cmdAgregar.Caption = "&Grabar"
    cmdCancelar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    LIMPIA_COM
    Txtcom(1).Locked = False
    Txtcom(0).Text = LK_CODCIA
    com_signo_d.ListIndex = 0
    com_signo_h.ListIndex = 2
    FrmComaest.com_tipo_cuenta.ListIndex = 0
    DESBLOQUEA_TEXT
    Modo.Caption = "< Agregar >"
    FMODO.Visible = True
    REPLACE1.Visible = True
    Txtcom(0).Enabled = False
    Txtcom(5).Enabled = False
    Txtcom(1).SetFocus
    'AGREGAMOS EN BLANCO
Else
   If FrmComaest.Txtcom(2).Text = "" Or Len(FrmComaest.Txtcom(2).Text) = 0 Then
       MsgBox "Ingrese Descripción  ..!!!", 48, Pub_Titulo
       Azul Txtcom(2), Txtcom(2)
       Exit Sub
   End If
   If Len(FrmComaest.Txtcom(1).Text) = 2 And FrmComaest.com_tipo_cuenta.ListIndex = 0 Then
       MsgBox "Cuenta Principal debe tener su Tipo   ..!!!", 48, Pub_Titulo
       com_tipo_cuenta.SetFocus
       Exit Sub
   End If
   PUB_CUENTA = Trim(FrmComaest.Txtcom(1).Text)
   SQ_OPER = 1
   LEER_COM_LLAVE
   If Not com_llave.EOF Then
     MsgBox "Cuenta Ya Existe ....", 48, Pub_Titulo
     cmdcancelar_Click
     REPLACE1.Visible = False
     Exit Sub
   End If
   VERI_NIVEL Trim(FrmComaest.Txtcom(1).Text)
   If NIVEL_ACT = 0 Then
     Exit Sub
   End If
   If NIVEL_ACT = NIVEL_MAX Then
     Txtcom(7).Text = "1"
   Else
     Txtcom(7).Text = "0"
   End If
   'SI EXISTEN LAS CUENTAS AUTOMATICAS ..
   If Trim(Txtcom(3).Text) <> "" Then
    If Not BUSCA_AUTO(Txtcom(3).Text) Then
      MsgBox "Cuenta Invalidad ..", 48, Pub_Titulo
      Azul Txtcom(3), Txtcom(3)
      Exit Sub
    End If
   End If
   If Trim(Txtcom(4).Text) <> "" Then
    If Not BUSCA_AUTO(Txtcom(4).Text) Then
      MsgBox "Cuenta Invalidad ..", 48, Pub_Titulo
      Azul Txtcom(4), Txtcom(4)
      Exit Sub
    End If
   End If
   If NIVEL_ACT <> 1 Then
     FrmComaest.com_tipo_cuenta.ListIndex = LOC_TIPO_CTA
   End If
   FrmComaest.Txtcom(6).Text = LOC_CTA_SUP
   FrmComaest.com_signo_d = LOC_SIG_DEB
   FrmComaest.com_signo_h = LOC_SIG_HAB
   Screen.MousePointer = 11
   GRABAR_COM
   MENSAJE_COM " AGREGADO... "
   'CmdAgregar.Caption = "&Agregar"
   'cmdCancelar.Enabled = True
   'CmdModificar.Enabled = True
   'cmdEliminar.Enabled = True
   LIMPIA_COM
   Txtcom(0).Text = LK_CODCIA
   com_signo_d.ListIndex = 0
   com_signo_h.ListIndex = 2
   com_tipo_cuenta.ListIndex = 0
   Txtcom(1).SetFocus
   'Txtcom(1).Locked = False
   'PASA = False
   'BLOQUEA_TEXT
   Screen.MousePointer = 0
      
End If
   
Exit Sub
SALE:
If Err.Number = 40002 Then
  MsgBox "Error ..al Grabar Cuenta .. llamar a Computo..", 48, Pub_Titulo
Else
 MsgBox Err.Description, 48, Pub_Titulo
End If
If cmdCancelar.Enabled Then
    cmdcancelar_Click
    Exit Sub
End If
Unload FrmComaest


End Sub

Private Sub cmdAgregar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Txtcom(1).SetFocus
End If

End Sub

Private Sub cmdcancelar_Click()
If Left(cmdAgregar.Caption, 2) = "&A" And Left(cmdModificar.Caption, 2) = "&M" Then
    LIMPIA_COM
    cmdCancelar.Enabled = False
    Txtcom(1).Locked = False
    MENSAJE_COM "Proceso Cancelado... !!!    "
    Modo.Caption = ""
    FMODO.Visible = False
    REPLACE1.Visible = False
    Txtcom(1).Enabled = True
    Txtcom(1).SetFocus
     Exit Sub
End If
    Screen.MousePointer = 11
     If Left(cmdModificar.Caption, 2) = "&G" Then
        cmdModificar.Caption = "&Modificar"
        LLENA_COM 1
        Txtcom(1).Locked = True
        BLOQUEA_TEXT
     Else
        cmdAgregar.Caption = "&Agregar"
        LIMPIA_COM
        Txtcom(1).Locked = False
        REPLACE1.Visible = False
        BLOQUEA_TEXT
     End If
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdModificar.Enabled = True
     cmdEliminar.Enabled = True
     Txtcom(1).Enabled = True
     Modo.Caption = ""
     FMODO.Visible = False
     MENSAJE_COM "Proceso Cancelado... !!!    "
     Txtcom(1).SetFocus
     Screen.MousePointer = 0
End Sub

Private Sub cmdcerrar_Click()
Unload FrmComaest

End Sub

Private Sub cmdCerrar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    FrmComaest.Txtcom(1).SetFocus
End If

End Sub

Private Sub cmdEliminar_Click()
Dim pos As Integer
Dim cad As String
Dim wcuenta As String
Dim WELI As rdoResultset
If Len(Txtcom(2).Text) = 0 Or Len(Txtcom(1).Text) = 0 Then
   MENSAJE_COM "NO a seleccionado NADA ... !"
   Exit Sub
End If
  pub_mensaje = " ¿Desea Eliminar el Registro... ?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbYes Then   ' El usuario eligió
    wcuenta = Trim(Txtcom(1).Text)
    cad = "SELECT * FROM COMAEST WHERE COM_CUENTA > '" & wcuenta & "' AND COM_CODCIA = '" & LK_CODCIA & "' ORDER BY COM_CUENTA"
    Set WELI = CN.OpenResultset(cad, rdOpenKeyset, rdConcurValues)
    WELI.Requery
    If Not WELI.EOF Then
       pos = Len(Txtcom(1).Text)
      If Left(WELI!com_cuenta, pos) = Trim(Txtcom(1).Text) Then
          MsgBox "No se puede Eliminar ..Existen SubCuentas...!!!", 48, Pub_Titulo
          Exit Sub
      End If
    End If
    Screen.MousePointer = 11
    'WELI.Delete
    com_llave.Delete
    FrmComaest.Txtcom(1).Text = ""
    FrmComaest.Txtcom(1).Locked = False
    LIMPIA_COM
    MENSAJE_COM "Registro, ELIMINADO ... "
    Screen.MousePointer = 0
   Exit Sub
  End If
  Screen.MousePointer = 0


End Sub

Private Sub CmdModificar_Click()
If Len(Txtcom(1)) = 0 Or Len(Txtcom(2)) = 0 Then
   MENSAJE_COM "NO a seleccionado NADA ... !"
   Exit Sub
End If
If Left(cmdModificar.Caption, 2) = "&M" Then
    cmdModificar.Caption = "&Grabar"
    cmdAgregar.Enabled = False
    cmdEliminar.Enabled = False
    cmdCancelar.Enabled = True
    Txtcom(1).Locked = True
    DESBLOQUEA_TEXT
    Txtcom(2).SetFocus
    Txtcom(0).Enabled = False
    Txtcom(5).Enabled = False
    Modo.Caption = "< Modificación >"
    FMODO.Visible = True
Else
    '*Grabar las modificaciones
    If Txtcom(2).Text = "" Or Len(Txtcom(2).Text) = 0 Then
         MsgBox " Nombre Invalido ....", 48, Pub_Titulo
         Exit Sub
    End If
   If Not Trim(Txtcom(4).Text) = "" Then
      If Not BUSCA_AUTO(Txtcom(3).Text) Then
        MsgBox "Cuenta NO Existe o  es Invalidad ..", 48, Pub_Titulo
        Azul Txtcom(3), Txtcom(3)
        Exit Sub
      End If
   End If
   If Not Trim(Txtcom(3).Text) = "" Then
     If Not BUSCA_AUTO(Txtcom(4).Text) Then
        MsgBox "Cuenta NO Existe o  es Invalidad ..", 48, Pub_Titulo
        Azul Txtcom(4), Txtcom(4)
        Exit Sub
      End If
   End If
     Screen.MousePointer = 11
     GRABAR_COM
     MENSAJE_COM "Registro , MODIFICADO... "
     cmdModificar.Caption = "&Modificar"
     cmdCancelar.Enabled = True
     cmdAgregar.Enabled = True
     cmdEliminar.Enabled = True
     Txtcom(1).Locked = False
     BLOQUEA_TEXT
     LIMPIA_COM
     Modo.Caption = ""
     FMODO.Visible = False
     Txtcom(1).SetFocus
     Screen.MousePointer = 0
    
End If

End Sub

Private Sub cmdModificar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    FrmComaest.Txtcom(1).SetFocus
End If

End Sub

Private Sub CmdRepo_Click()
Dim RUTA As String, wconect, wname
RUTA = PUB_RUTA_REPORTE & "cuentas.rpt"
wconect = PUB_ODBC '& "ODBC;DRIVER={SQL Server};UID=" & CONTS_UID & ";PWD=" & CONTS_PWD & ";APP=Data From Designer;WSID=" & CONTS_SERVER & ";LANGUAGE=us_english;DATABASE=BDATOS"
wname = ""
'ListaRepo.UserName = wname
ListaRepo.Connect = wconect
ListaRepo.ReportFileName = RUTA
ListaRepo.WindowTitle = "Lista  Cuentas ... " '& Trim(Left(CmbBancos, 20))
ListaRepo.SelectionFormula = "{COMAEST.COM_CODCIA} = '" & LK_CODCIA & "'"
ListaRepo.WindowLeft = 2
ListaRepo.WindowTop = 70
ListaRepo.WindowWidth = 635
ListaRepo.WindowHeight = 390
DoEvents
ListaRepo.Action = 1

'{COMAEST.COM_CODCIA} = "01"
End Sub

Private Sub com_signo_d_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  com_signo_h.SetFocus
End If
End Sub

Private Sub com_signo_h_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If cmdAgregar.Enabled Then
       cmdAgregar.SetFocus
  ElseIf cmdModificar.Enabled Then
       cmdModificar.SetFocus
  End If
End If

End Sub

Private Sub com_tipo_cuenta_Click()
If Left(cmdAgregar.Caption, 2) <> "&G" Then
 Exit Sub
End If
If com_tipo_cuenta.ListIndex = 2 Or com_tipo_cuenta.ListIndex = 4 Then
 com_signo_d.ListIndex = 2
 com_signo_h.ListIndex = 0
Else
 com_signo_d.ListIndex = 0
 com_signo_h.ListIndex = 2
End If
End Sub

Private Sub com_tipo_cuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If cmdAgregar.Enabled Then
    cmdAgregar.SetFocus
  ElseIf cmdModificar.Enabled Then
    cmdModificar.SetFocus
  End If
End If
End Sub

Private Sub Command1_Click()
Dim W As Integer
Dim cta1 As String
Dim cta2 As String
Dim cta3 As Integer
Dim signo_d As Integer
Dim signo_h As Integer
Dim una

cta1 = Left(Txtcom(1).Text, 2)
cta2 = Left(Txtcom(2).Text, 2)
signo_d = Val(com_signo_d.Text)
signo_h = Val(com_signo_h.Text)
If Len(cta1) = 0 And Len(cta1) = 0 Then
  GoTo fin
End If
If signo_h = 0 And signo_d = 0 Then
  GoTo fin
End If
SQ_OPER = 1
PUB_CUENTA = cta1
LEER_COM_LLAVE
If com_llave.EOF Then
   MsgBox "Cuenta ,   NO EXISTE ... ", 48, Pub_Titulo
   Azul FrmComaest.Txtcom(1), FrmComaest.Txtcom(1)
   GoTo fin
End If
SQ_OPER = 1
PUB_CUENTA = cta2
LEER_COM_LLAVE
If com_llave.EOF Then
   MsgBox "Cuenta ,   NO EXISTE ... ", 48, Pub_Titulo
   Azul FrmComaest.Txtcom(2), FrmComaest.Txtcom(2)
   GoTo fin
End If
Dim PC_CUENTA As rdoQuery
Dim ps_cta As rdoResultset

cad = "SELECT * FROM COMAEST WHERE COM_CUENTA >= ?  AND COM_CUENTA <= ? and COM_CODCIA = ? ORDER BY COM_CUENTA"
Set PC_CUENTA = CN.CreateQuery("", cad)
Set ps_cta = PC_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)

'If cta1 = cta2 Then
 cta3 = Val(cta2) + 1
 cta2 = cta3
'End If

PC_CUENTA(0) = cta1
PC_CUENTA(1) = cta2
PC_CUENTA(2) = LK_CODCIA
ps_cta.Requery
ProgBar.Min = 0
ProgBar.max = ps_cta.RowCount
Command1.Enabled = False
ProgBar.Visible = True
ProgBar.Value = 0
W = 0
ps_cta.MoveFirst
Do Until ps_cta.EOF
   'If cta1 = cta2 Then
   
   a = ps_cta!com_cuenta  '  Exit Do
   'End If
   ps_cta.Edit
   ps_cta!com_signo_d = signo_d
   ps_cta!com_signo_h = signo_h
   ps_cta.Update
   W = W + 1
   ProgBar.Value = W
ps_cta.MoveNext
Loop
ProgBar.Visible = False
Command1.Enabled = True
MsgBox "Proceso Terminado", 48, Pub_Titulo
fin:
End Sub

Private Sub Form_Load()
UNICO = " "
Dim cade As String
Wflag = "0"
IND = 0
CU = 0
pasa = False
LIMPIA_COM
com_signo_h.AddItem "1"
com_signo_h.AddItem "0"
com_signo_h.AddItem "-1"
com_signo_d.AddItem "1"
com_signo_d.AddItem "0"
com_signo_d.AddItem "-1"

FrmComaest.com_tipo_cuenta.Clear
PUB_TIPREG = 16
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
  FrmComaest.com_tipo_cuenta.AddItem Nulo_Valors(tab_mayor!TAB_NOMCORTO) & String(20, " ") & tab_mayor!TAB_NUMTAB
  tab_mayor.MoveNext
Loop

BLOQUEA_TEXT
cade = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", cade)
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PSCOP_LLAVE.rdoParameters(0) = par_llave!PAR_CIACON
cop_llave.Requery
If Not cop_llave.EOF Then
For i = 1 To 6
  If cop_llave.rdoColumns(i) <> 0 Then
     wCOM_NIVEL(i) = cop_llave.rdoColumns(i)
     NIVEL_MAX = i
  End If
Next i
End If



End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
     FLISTA.Visible = False
     If FrmComaest.Txtcom(IND).Enabled Then
         FrmComaest.Txtcom(IND).SetFocus
         Exit Sub
     End If
     FrmComaest.Txtcom(1).SetFocus
     IND = 0

End If

End Sub

Private Sub GRID1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
     FLISTA.Visible = False
     FrmComaest.Txtcom(1).SetFocus
     KeyCode = 0
     Exit Sub
End If
If KeyCode = 13 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) = "" Then
      Exit Sub
    End If
    If IND = 3 Then
      If cmdAgregar.Enabled And cmdModificar.Enabled Then
        GoTo sigue
      End If
      grid1.COL = 2
      If grid1.Text <> "1" Then
         MsgBox "No Procede .. Cuenta debe ser inferior ..", 48, Pub_Titulo
         FLISTA.Visible = False
         Txtcom(3).SetFocus
         IND = 0
         Exit Sub
      End If
      grid1.COL = 0
      Txtcom(3).Text = grid1.Text
      FLISTA.Visible = False
      Txtcom(3).SetFocus
      IND = 0
      Exit Sub
    End If
    If IND = 4 Then
      If cmdAgregar.Enabled And cmdModificar.Enabled Then
        GoTo sigue
      End If
      grid1.COL = 2
      If grid1.Text <> "1" Then
         MsgBox "No Procede .. Cuenta debe ser inferior ..", 48, Pub_Titulo
         FLISTA.Visible = False
         Txtcom(4).SetFocus
         grid1.COL = 0
         IND = 0
         Exit Sub
      End If
      grid1.COL = 0
      Txtcom(4).Text = grid1.Text
      FLISTA.Visible = False
      Txtcom(4).SetFocus
      IND = 0
      Exit Sub
    End If
sigue:
    If Left(grid1.Text, 1) = " " Or Len(grid1.Text) = 0 Then
       Exit Sub
    End If
    LLENA_COM 0
    FLISTA.Visible = False
    FrmComaest.Txtcom(1).Locked = True
    cmdCancelar.Enabled = True
    cmdModificar.SetFocus
    KeyCode = 0
End If

End Sub

Private Sub grid1_LostFocus()
'FLISTA.Visible = False
'FrmComaest.Txtcom(1).SetFocus
End Sub

Private Sub LEIDO_Change()
Dim WSELE As String
'DoEvents
PRO_CUENTAS "X"
'grid1.Col = 0
'grid1.Row = 1
'grid1.ColSel = grid1.Cols - 1
'grid1.RowSel = 1

End Sub

Private Sub LEIDO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
     FLISTA.Visible = False
     If FrmComaest.Txtcom(IND).Enabled Then
         FrmComaest.Txtcom(IND).SetFocus
         Exit Sub
     End If
     FrmComaest.Txtcom(1).SetFocus
     IND = 0
End If

End Sub

Private Sub LEIDO_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Then
    grid1.COL = 0
    grid1.Row = 1
    grid1.SetFocus
Else

End If

End Sub

Private Sub PARPADEA_Timer()
 CU = CU + 1
 LblMensaje.Visible = True 'Not LblMensaje.Visible
 If CU > 2 Then
   CU = 0
   Parpadea.Enabled = False
   LblMensaje.Visible = False
 End If

End Sub

Private Sub Txtcom_Change(Index As Integer)
If Index = 1 And Left(cmdAgregar.Caption, 2) = "&G" Then
    If Len(Txtcom(1).Text) = 2 Then
      Txtcom(3).Text = ""
      Txtcom(4).Text = ""
      Txtcom(3).Enabled = False
      Txtcom(4).Enabled = False
      com_tipo_cuenta.Enabled = True
      com_signo_d.Enabled = True
      com_signo_h.Enabled = True
    Else
     com_tipo_cuenta.Enabled = False
     Txtcom(3).Enabled = True
     Txtcom(4).Enabled = True
     com_signo_d.Enabled = False
     com_signo_h.Enabled = False
     com_signo_d.ListIndex = 0
     com_signo_h.ListIndex = 2
     com_tipo_cuenta.ListIndex = 0
    End If
End If
End Sub

Private Sub Txtcom_KeyPress(Index As Integer, KeyAscii As Integer)
Static ca As Integer
Dim WSELE As String
Dim wnivel
Select Case Index
Case 4
    'If Txtcom(4).text <> "" Then
      If KeyAscii = 13 Then
        If Len(Trim(Txtcom(1).Text)) = wCOM_NIVEL(1) Then
            FrmComaest.com_tipo_cuenta.SetFocus
            SendKeys "%{UP}"
        Else
           If cmdAgregar.Enabled Then
             cmdAgregar.SetFocus
           ElseIf cmdModificar.Enabled Then
             cmdModificar.SetFocus
           End If
        End If
       End If
      Exit Sub
    'End If
    If KeyAscii = 13 Then
      FrmComaest.com_signo_h.SetFocus
    End If
Case 3
    'If Txtcom(3).text <> "" Then
    ' If KeyAscii = 13 Then
    '   Txtcom(4).SetFocus
    ' End If
    '  Exit Sub
    'End If
    If KeyAscii = 13 Then
     Txtcom(4).SetFocus
    ' FrmComaest.com_signo_d.SetFocus
    End If

Case 2
   If KeyAscii = 13 Then
      If Left(Txtcom(1).Text, 1) = "6" Then
        TxtCtaSobre.SetFocus
        Exit Sub
       End If
      If com_tipo_cuenta.Enabled Then
         com_tipo_cuenta.SetFocus
      Else
        If cmdAgregar.Enabled Then
            cmdAgregar.SetFocus
        Else
         cmdModificar.SetFocus
        End If
      End If
      Exit Sub
   End If
    
Case 1
    If KeyAscii <> 13 Then
     Exit Sub
    End If
    If Left(FrmComaest.cmdAgregar.Caption, 2) = "&G" Or Left(FrmComaest.cmdModificar.Caption, 2) = "&G" Then
        Txtcom(2).SetFocus
        Exit Sub
    End If
    Dim valor As String
    Dim tf As Integer
    If FrmComaest.Txtcom(1).Text = "" Then
        'wsele = "SELECT * FROM COMAEST WHERE COM_CUENTA >= ? ORDER BY COM_CUENTA"
       WSELE = "SELECT * FROM COMAEST WHERE COM_CUENTA >= ? AND COM_CODCIA = '" & LK_CODCIA & "' ORDER BY COM_CUENTA"
       LEIDO.Text = ""
       llave1 = ""
       PRO_CUENTAS WSELE
       FLISTA.Visible = True
       LEIDO.SetFocus
    Else
        SQ_OPER = 1
        PUB_CUENTA = FrmComaest.Txtcom(1).Text
        LEER_COM_LLAVE
        If com_llave.EOF Then
                MENSAJE_COM "Registro ,   NO EXISTE ... "
                Azul FrmComaest.Txtcom(1), FrmComaest.Txtcom(1)
                GoTo fin
        End If
        LLENA_COM 1
        FrmComaest.Txtcom(1).Locked = True
        cmdCancelar.Enabled = True
        cmdModificar.SetFocus
   End If

End Select

fin:
End Sub

Public Sub LISTA_CUENTAS()
grid1.Cols = 3
grid1.COL = 0
grid1.Row = 0
grid1.ColAlignment(0) = 1
grid1.ColAlignment(1) = 1
grid1.ColAlignment(2) = 1

grid1.Text = " Cuenta "
grid1.COL = 1
grid1.Text = " Descripción "
grid1.ColWidth(0) = 1300
grid1.ColWidth(1) = 2300
grid1.ColWidth(2) = 1

End Sub
Public Function PRO_CUENTAS(archi As String)
Dim NUMCAMPO As Integer
Dim OJO As String * 1

wBusca = "SELECT * FROM COMAEST WHERE COM_CUENTA >= ? AND COM_CODCIA = ? ORDER BY COM_CUENTA"
If UNICO <> wBusca Then
   Set PSX = CN.CreateQuery("", wBusca)
End If

PSX.rdoParameters(0) = LEIDO.Text
PSX.rdoParameters(1) = LK_CODCIA
If UNICO <> wBusca Then
   Set X = PSX.OpenResultset(rdOpenKeyset)
End If
   
If UNICO = wBusca Then
   X.Requery
   If X.RowCount > 0 Then
      X.MoveFirst
   End If
End If
If X.EOF = True Then
   grid1.Clear
   grid1.Visible = True
   grid1.Row = 1
   grid1.COL = 1
   grid1.Text = "No hay registros"
   llave1 = ""
   UNICO = wBusca
   'LEIDO.Locked = True
   Exit Function
End If

If X.rdoColumns(1) = llave1 Then
    DoEvents
    Exit Function
End If
If X.RowCount > 0 Then
   llave1 = X.rdoColumns(1)
End If

UNICO = wBusca
grid1.Rows = 2
fila = 0
grid1.Clear
LISTA_CUENTAS
LEIDO.Locked = False
'grid1.Rows = 2
'grid1.Cols = 4

fila = 0
Do Until X.EOF Or fila = 100
    fila = fila + 1
    grid1.TextMatrix(fila, 0) = Nulo_Valors(X.rdoColumns(1))
    grid1.TextMatrix(fila, 1) = Nulo_Valors(X.rdoColumns(2))
    grid1.TextMatrix(fila, 2) = X.rdoColumns(7)
    X.MoveNext
    grid1.Rows = grid1.Rows + 1
Loop
grid1.TextMatrix(fila + 1, 1) = "                * * *    END    * * * "
grid1.Row = 1
grid1.COL = 1
grid1.Visible = True
End Function


Public Sub BLOQUEA_TEXT()
Txtcom(0).Enabled = False
Txtcom(2).Enabled = False
Txtcom(3).Enabled = False
Txtcom(4).Enabled = False
Txtcom(5).Enabled = False
TxtCtaSobre.Enabled = False
com_tipo_cuenta.Enabled = False
com_signo_d.Enabled = False
com_signo_h.Enabled = False




End Sub
Public Sub DESBLOQUEA_TEXT()
Txtcom(0).Enabled = True
Txtcom(2).Enabled = True
Txtcom(3).Enabled = True
Txtcom(4).Enabled = True
Txtcom(5).Enabled = True
TxtCtaSobre.Enabled = True
com_tipo_cuenta.Enabled = True
com_signo_d.Enabled = True
com_signo_h.Enabled = True

End Sub


Public Sub VERI_NIVEL(valor As String)
Dim cade As String
Dim wnivel As Integer
Dim i, windex As Integer
Dim wBusca As String
LOC_CTA_SUP = ""
NIVEL_ACT = 0
cade = Trim(valor)
wnivel = Len(cade)
windex = 0
For i = 1 To 6
  If wCOM_NIVEL(i) = wnivel Then
     windex = i
     Exit For
  End If
Next i
If windex = 0 Then
    MsgBox "Nro. de digitos invalidos, Verificar ..", 48, Pub_Titulo
    NIVEL_ACT = 0
   Exit Sub
ElseIf windex <> 1 Then
   windex = windex - 1
Else
   'msgbox "Es nivel Superior  .."
   NIVEL_ACT = 1
   Exit Sub
End If
wBusca = Left(cade, wCOM_NIVEL(windex))
SQ_OPER = 1
PUB_CUENTA = wBusca
LEER_COM_LLAVE
If com_llave.EOF Then
   MsgBox "Cuenta Superior NO Existe, Verificar ...", 48, Pub_Titulo
   Exit Sub
End If
LOC_TIPO_CTA = com_llave!COM_TIPO_CTA
LOC_CTA_SUP = com_llave!com_cuenta
LOC_SIG_DEB = com_llave!com_signo_d
LOC_SIG_HAB = com_llave!com_signo_h

NIVEL_ACT = windex + 1
End Sub

Private Sub Txtcom_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim CADE3
If Index = 4 Or Index = 3 Then
  If KeyCode = 46 Then
    Txtcom(Index).Text = " "
  End If
  If KeyCode = 45 Then
   '  If KeyAscii = 13 Then
       archi = "SELECT * FROM COMAEST WHERE COM_DESCRIPCION >= ? AND COM_CODCIA = '" & LK_CODCIA & "' ORDER BY COM_DESCRIPCION"
       llave1 = ""
       LEIDO.Text = ""
       PRO_CUENTAS archi
       FLISTA.Visible = True
       LEIDO.SetFocus
       IND = Index
    'End If
  End If
End If
End Sub

Public Function BUSCA_AUTO(valor As String) As Boolean
Dim ba As rdoResultset
Dim cade As String
  cade = "SELECT * FROM COMAEST WHERE COM_CUENTA = '" & Trim(valor) & "' AND COM_CODCIA = '" & LK_CODCIA & "' ORDER BY COM_CUENTA"
  Set ba = CN.OpenResultset(cade, rdOpenKeyset, rdConcurValues)
   ba.Requery
   BUSCA_AUTO = False
   If ba.EOF Then
     BUSCA_AUTO = False
     Exit Function
   End If
   If ba!COM_FLAG_AFECTACION = "1" Then
     BUSCA_AUTO = True
   End If

End Function

Private Sub TxtCtaSobre_KeyPress(KeyAscii As Integer)
Dim car
If KeyAscii = 13 Then
  If com_tipo_cuenta.Enabled Then
    com_tipo_cuenta.SetFocus
  Else
   If cmdAgregar.Enabled Then
      cmdAgregar.SetFocus
   Else
      cmdModificar.SetFocus
   End If
 End If
End If
car = Chr$(KeyAscii)
car = UCase$(Chr$(KeyAscii))
If car = "S" Or car = "N" Or KeyAscii = 8 Then
  KeyAscii = Asc(car)
  Exit Sub
End If
KeyAscii = 0
Beep
Exit Sub

End Sub
