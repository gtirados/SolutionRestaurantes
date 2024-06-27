VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCoPar 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Definición de Estructura en Contabilidad"
   ClientHeight    =   5820
   ClientLeft      =   1290
   ClientTop       =   1680
   ClientWidth     =   9915
   Icon            =   "FrmCoPar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   9915
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Digitos Centro Costos"
      ForeColor       =   &H00800000&
      Height          =   2115
      Left            =   8055
      TabIndex        =   44
      Top             =   3705
      Width           =   1815
      Begin VB.TextBox txtdgcc 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   630
         TabIndex        =   38
         Text            =   "1"
         Top             =   1740
         Width           =   600
      End
      Begin VB.TextBox txtdgcc 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   630
         TabIndex        =   36
         Text            =   "1"
         Top             =   1125
         Width           =   600
      End
      Begin VB.TextBox txtdgcc 
         Height          =   285
         Index           =   0
         Left            =   630
         TabIndex        =   34
         Text            =   "1"
         Top             =   540
         Width           =   600
      End
      Begin VB.CheckBox chkcc 
         BackColor       =   &H00FAEFDA&
         Caption         =   "CC3"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   195
         TabIndex        =   37
         Top             =   1470
         Width           =   1290
      End
      Begin VB.CheckBox chkcc 
         BackColor       =   &H00FAEFDA&
         Caption         =   "CC2"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   35
         Top             =   870
         Width           =   1290
      End
      Begin VB.CheckBox chkcc 
         BackColor       =   &H00FAEFDA&
         Caption         =   "CC1"
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   33
         Top             =   285
         Value           =   1  'Checked
         Width           =   1290
      End
   End
   Begin VB.Frame FRACLAVE 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Clave de Acceso  :"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   5880
      TabIndex        =   28
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox txtpas 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         MaxLength       =   8
         PasswordChar    =   "*"
         TabIndex        =   29
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdgrabar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   15
      Top             =   5355
      Width           =   1575
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6435
      TabIndex        =   14
      Top             =   5355
      Width           =   1455
   End
   Begin VB.Frame fraconta 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Opciones Contables"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5295
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   7815
      Begin VB.CheckBox cheact 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Asignar Periodo Activo."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   3855
      End
      Begin VB.ComboBox ano 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "FrmCoPar.frx":0442
         Left            =   4320
         List            =   "FrmCoPar.frx":0467
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdest 
         Height          =   375
         Left            =   5760
         TabIndex        =   27
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox MES 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         ItemData        =   "FrmCoPar.frx":04AD
         Left            =   855
         List            =   "FrmCoPar.frx":04AF
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   2775
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3135
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5530
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         BackColor       =   16445402
         ForeColor       =   12582912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Estado Contable"
         TabPicture(0)   =   "FrmCoPar.frx":04B1
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ListCias"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Tit"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame2(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Parametros Contables"
         TabPicture(1)   =   "FrmCoPar.frx":04CD
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblfac"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame2(1)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin VB.Frame Frame2 
            BackColor       =   &H00FAEFDA&
            Height          =   2790
            Index           =   1
            Left            =   -74955
            TabIndex        =   46
            Top             =   315
            Width           =   7140
            Begin VB.TextBox c_tcg 
               Height          =   285
               Left            =   2565
               MaxLength       =   12
               TabIndex        =   56
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox c_tcp 
               Height          =   285
               Left            =   2565
               MaxLength       =   12
               TabIndex        =   55
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox c_redg 
               Height          =   285
               Left            =   2565
               MaxLength       =   12
               TabIndex        =   54
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox c_redp 
               Height          =   285
               Left            =   2565
               MaxLength       =   12
               TabIndex        =   53
               Top             =   2040
               Width           =   1335
            End
            Begin VB.Label lblc 
               BackColor       =   &H00FAEFDA&
               BackStyle       =   0  'Transparent
               Caption         =   "Cta.  x Diferencia de T.C. (Ganancia) :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   435
               Left            =   225
               TabIndex        =   60
               Top             =   360
               Width           =   1935
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FAEFDA&
               BackStyle       =   0  'Transparent
               Caption         =   "Cta.  x Diferencia de T.C. (Perdida) :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   435
               Left            =   225
               TabIndex        =   59
               Top             =   840
               Width           =   1785
            End
            Begin VB.Label Label3 
               BackColor       =   &H00FAEFDA&
               BackStyle       =   0  'Transparent
               Caption         =   "Cta.  x Diferencia de Redondeo (Ganancia) :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   435
               Left            =   225
               TabIndex        =   58
               Top             =   1560
               Width           =   2385
            End
            Begin VB.Label Label4 
               BackColor       =   &H00FAEFDA&
               BackStyle       =   0  'Transparent
               Caption         =   "Cta.  x Diferencia de Redondeo (Perdida) :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   435
               Left            =   225
               TabIndex        =   57
               Top             =   2040
               Width           =   2235
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FAEFDA&
            Height          =   2790
            Index           =   0
            Left            =   45
            TabIndex        =   45
            Top             =   315
            Width           =   7140
            Begin VB.Shape Che_pase 
               BorderColor     =   &H00000000&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   420
               Shape           =   3  'Circle
               Top             =   375
               Width           =   375
            End
            Begin VB.Shape Che_regc 
               BorderColor     =   &H00000000&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   420
               Shape           =   3  'Circle
               Top             =   735
               Width           =   375
            End
            Begin VB.Shape che_regv 
               BorderColor     =   &H00000000&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   420
               Shape           =   3  'Circle
               Top             =   1095
               Width           =   375
            End
            Begin VB.Shape Che_caja 
               BorderColor     =   &H00000000&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   420
               Shape           =   3  'Circle
               Top             =   1455
               Width           =   375
            End
            Begin VB.Shape Che_destinos 
               BorderColor     =   &H00000000&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   420
               Shape           =   3  'Circle
               Top             =   1815
               Width           =   375
            End
            Begin VB.Shape Che_mayor 
               BorderColor     =   &H00000000&
               FillColor       =   &H000000FF&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   420
               Shape           =   3  'Circle
               Top             =   2175
               Width           =   375
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pase de Contabilidad"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   3
               Left            =   900
               TabIndex        =   52
               Top             =   375
               Width           =   1755
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Registro de Compra"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   4
               Left            =   900
               TabIndex        =   51
               Top             =   735
               Width           =   1680
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Registro de Venta"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   5
               Left            =   900
               TabIndex        =   50
               Top             =   1095
               Width           =   1515
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Registro de Caja y Bancos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   900
               TabIndex        =   49
               Top             =   1455
               Width           =   2190
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Proceso x Ctas de Destinos"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   7
               Left            =   900
               TabIndex        =   48
               Top             =   1815
               Width           =   2280
            End
            Begin VB.Label lbl1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Proceso de Mayorización"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   8
               Left            =   900
               TabIndex        =   47
               Top             =   2175
               Width           =   2100
            End
         End
         Begin VB.ListBox Tit 
            Height          =   2535
            Left            =   7320
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ListBox ListCias 
            Height          =   1860
            Left            =   7560
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   1920
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Label lblfac 
            Height          =   255
            Left            =   -74760
            TabIndex        =   24
            Top             =   600
            Width           =   2775
            WordWrap        =   -1  'True
         End
      End
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCampo1 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label act_ano 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   4680
         TabIndex        =   43
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label act_mes 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   3840
         TabIndex        =   42
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mes.         Año."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3840
         TabIndex        =   41
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   840
         TabIndex        =   40
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label estado 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Año:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl1 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo Contable Activo."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Width           =   2295
      End
   End
   Begin VB.Frame franivel 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Digitos para el nivel"
      ForeColor       =   &H00800000&
      Height          =   3615
      Left            =   8040
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      Begin VB.TextBox txtnivelmax 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         MaxLength       =   1
         TabIndex        =   32
         Top             =   3135
         Width           =   720
      End
      Begin VB.TextBox txtCopar 
         Height          =   285
         Index           =   5
         Left            =   960
         MaxLength       =   2
         TabIndex        =   6
         Top             =   2460
         Width           =   720
      End
      Begin VB.TextBox txtCopar 
         Height          =   285
         Index           =   4
         Left            =   960
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2016
         Width           =   720
      End
      Begin VB.TextBox txtCopar 
         Height          =   285
         Index           =   3
         Left            =   960
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1572
         Width           =   720
      End
      Begin VB.TextBox txtCopar 
         Height          =   285
         Index           =   2
         Left            =   960
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1128
         Width           =   720
      End
      Begin VB.TextBox txtCopar 
         Height          =   285
         Index           =   1
         Left            =   960
         MaxLength       =   2
         TabIndex        =   2
         Top             =   684
         Width           =   720
      End
      Begin VB.TextBox txtCopar 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Maximo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   2895
         Width           =   1035
      End
      Begin VB.Label lblcopar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 6 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   2460
         Width           =   585
      End
      Begin VB.Label lblcopar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 5 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   2010
         Width           =   585
      End
      Begin VB.Label lblcopar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 4 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   1575
         Width           =   585
      End
      Begin VB.Label lblcopar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 3 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label lblcopar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 2 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   690
         Width           =   585
      End
      Begin VB.Label lblcopar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel 1 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   585
      End
   End
End
Attribute VB_Name = "FrmCoPar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REP_FECHA1
Dim REP_FECHA2

Option Explicit

Private Sub cmdcerrar_Click()
Unload FrmCoPar
End Sub

Private Sub cmdest_Click()
Dim WPASS As String
If Left(cmdest.Caption, 1) = "C" Then
    pub_mensaje = "Desea Cerrar el Mes de : " & Trim(MES.Text) & " Continuar..? "
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
        Exit Sub
    End If

    OEMES Left(MES.Text, 2), "01/01/" & Trim(ano.Text), 1
    Exit Sub
End If
FRACLAVE.Visible = True
txtpas.SetFocus
End Sub

Private Sub cmdgrabar_Click()
Dim wfecha1
Dim wfecha2
Dim WDIAS As Integer
Dim wmes As Integer
Dim DG_CC As String
Dim i As Integer

If Not SON_FECHAS(txtCampo1, txtCampo2) Then
   Exit Sub
End If
If Val(txtnivelmax.Text) <= 0 Then
 MsgBox "Ingrese el Nivel Maximo...", 48, Pub_Titulo
 txtnivelmax.SetFocus
 Exit Sub
End If
For fila = 1 To Val(txtnivelmax.Text)
 If txtCopar(fila - 1).Text = "" Then
  MsgBox "Falta ingresar algun Nivel...", 48, Pub_Titulo
  txtCopar(fila - 1).SetFocus
  Exit Sub
 End If
Next fila

For fila = 1 To 6
 If txtCopar(fila - 1).Text <> "" Then
   If fila = 1 And Val(txtCopar(fila - 1).Text) > 0 Then
    GoTo SIGUE
   ElseIf fila = 1 Then
     MsgBox "No Procede ..", 48, Pub_Titulo
     Exit Sub
   End If
   If Val(txtCopar(fila - 1).Text) < Val(txtCopar(fila - 2).Text) Then
     MsgBox "No Procede ..", 48, Pub_Titulo
     Exit Sub
   End If
 End If
SIGUE:
Next fila

SQ_OPER = 1
If Trim(c_tcg.Text) <> "" Then
    PUB_CUENTA = Trim(c_tcg.Text)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      MsgBox "Cuenta No Existe .  Intente Nuevamente . . .", 48, Pub_Titulo
      Azul c_tcg, c_tcg
      Exit Sub
    Else
      If Val(com_llave!com_nivel) <> Val(txtnivelmax.Text) Then
         MsgBox "Cuenta debe ser del Ultimo Nivel . . .", 48, Pub_Titulo
         Azul c_tcg, c_tcg
         Exit Sub
      End If
    End If
End If
If Trim(c_tcp.Text) <> "" Then
    PUB_CUENTA = Trim(c_tcp.Text)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      MsgBox "Cuenta No Existe .  Intente Nuevamente . . .", 48, Pub_Titulo
      Azul c_tcp, c_tcp
      Exit Sub
    Else
      If Val(com_llave!com_nivel) <> Val(txtnivelmax.Text) Then
         MsgBox "Cuenta debe ser del Ultimo Nivel . . .", 48, Pub_Titulo
         Azul c_tcp, c_tcp
         Exit Sub
      End If
    End If
End If
If Trim(c_redg.Text) <> "" Then
    PUB_CUENTA = Trim(c_redg.Text)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      MsgBox "Cuenta No Existe .  Intente Nuevamente . . .", 48, Pub_Titulo
      Azul c_redg, c_redg
      Exit Sub
    Else
      If Val(com_llave!com_nivel) <> Val(txtnivelmax.Text) Then
         MsgBox "Cuenta debe ser del Ultimo Nivel . . .", 48, Pub_Titulo
         Azul c_redg, c_redg
         Exit Sub
      End If
    End If
End If
If Trim(c_redp.Text) <> "" Then
    PUB_CUENTA = Trim(c_redp.Text)
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
      MsgBox "Cuenta No Existe .  Intente Nuevamente . . .", 48, Pub_Titulo
      Azul c_redp, c_redp
      Exit Sub
    Else
      If Val(com_llave!com_nivel) <> Val(txtnivelmax.Text) Then
         MsgBox "Cuenta debe ser del Ultimo Nivel . . .", 48, Pub_Titulo
         Azul c_redp, c_redp
         Exit Sub
      End If
    End If
End If

WDIAS = dias_mes(Val(Left(MES.Text, 2)), Val(ano.Text))
wmes = Val(Left(MES.Text, 2))
If wmes = 0 Then
 wfecha1 = "01" & "/" & Format(wmes + 1, "00") & "/" & Trim(ano.Text)
 wfecha2 = Format(1, "00") & "/" & Format(wmes + 1, "00") & "/" & Trim(ano.Text)
Else
 wfecha1 = "01" & "/" & Left(MES.Text, 2) & "/" & Trim(ano.Text)
 wfecha2 = Format(WDIAS, "00") & "/" & Left(MES.Text, 2) & "/" & Trim(ano.Text)
End If

REP_FECHA1 = wfecha1
REP_FECHA2 = wfecha2

If cop_llave.EOF Then
  cop_llave.AddNew
  cop_llave!COP_CODCIA = LK_CODCIA
Else
  cop_llave.Edit
End If
For fila = 1 To 6
 If txtCopar(fila - 1).Text <> "" Then
    cop_llave.rdoColumns(fila) = txtCopar(fila - 1).Text
 Else
    cop_llave.rdoColumns(fila) = 0
 End If
Next fila
If cheact.Value = 1 Then
 cop_llave!cop_fecha_proceso = REP_FECHA1
 cop_llave!COP_FECHA_PROCESO2 = REP_FECHA2
 cop_llave!cop_nro_mes = wmes
End If
cop_llave!cop_nivel_max = Val(txtnivelmax.Text)
cop_llave!COP_CTA_DIF_TC_favor = c_tcg.Text
cop_llave!COP_CTA_DIF_TC_CONTRA = c_tcp.Text
cop_llave!cop_cta_red_FAVOR = c_redg.Text
cop_llave!cop_cta_red_contra = c_redp.Text

DG_CC = ""
For i = 0 To 2
    If chkcc(i).Value = 1 Then
        DG_CC = DG_CC & "1"
        DG_CC = DG_CC & txtdgcc(i).Text
    ElseIf chkcc(i).Value = 0 Then
        DG_CC = DG_CC & "0"
        DG_CC = DG_CC & "0"
    End If
Next i
cop_llave!cop_flag_cc = DG_CC
cop_llave.Update
LK_NRO_MES = wmes
LK_FECHA_COP1 = REP_FECHA1
LK_FECHA_COP2 = REP_FECHA2
If LK_NRO_MES = 0 Then
   MDIForm1.tperiodo.Text = "Periodo: " & "APERTURA - " & Format(LK_FECHA_COP1, "yyyy")
 Else
   MDIForm1.tperiodo.Text = "Periodo: " & Format(LK_FECHA_COP1, "mmmm yyyy")
 End If
pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER>=? AND COV_FECHA_VOUCHER <=? AND COV_NRO_MES = " & LK_NRO_MES & "  ORDER BY COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSCOV_VOUCHER = CN.CreateQuery("", pub_cadena)
PSCOV_VOUCHER(0) = 0
PSCOV_VOUCHER(1) = LK_FECHA_DIA
PSCOV_VOUCHER(2) = LK_FECHA_DIA
Set cov_voucher = PSCOV_VOUCHER.OpenResultset(rdOpenKeyset, rdConcurValues)

MsgBox "Parametros de la Estructura a sido Modificado..", 48, Pub_Titulo
Dim iFormCount As Integer
On Error GoTo SIGUE
Screen.MousePointer = 11
If Forms.Count - 1 > 0 Then
   For iFormCount = Forms.Count - 1 To 1 Step -1
    If iFormCount <> 1 Then
        Unload Forms(iFormCount)
    End If
   Next iFormCount
End If
Unload FrmCoPar
Screen.MousePointer = 0
Exit Sub
End Sub

Private Sub Form_Load()
CenterMe FrmCoPar
Dim PSVER As rdoQuery
Dim ver_cuentas As rdoResultset
Dim tmp As String

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? "
Set PSVER = CN.CreateQuery("", pub_cadena)
PSVER(0) = 0
PSVER.MaxRows = 1
Set ver_cuentas = PSVER.OpenResultset(rdOpenKeyset, rdConcurValues)
PSVER.rdoParameters(0) = LK_CODCIA
ver_cuentas.Requery
LLENA_FECHAS
If Not ver_cuentas.EOF Then
'  MsgBox "Existe Cuenta en Cia. si desea Corregir la estructura, debera Eliminar todo el plan de Cuentas ", 48, Pub_Titulo
End If
cop_llave.Requery
If cop_llave.EOF Then
  MsgBox "Crear Registro Inicial para Evitar Efectos en el Sistema", 48, Pub_Titulo
  GoTo fin
Else
 For fila = 1 To 6
  If cop_llave.rdoColumns(fila) <> 0 Then
     txtCopar(fila - 1).Text = cop_llave.rdoColumns(fila)
  Else
    Exit For
  End If
 Next fila
 c_tcg.Text = Trim(cop_llave!COP_CTA_DIF_TC_favor)
 c_tcp.Text = Trim(cop_llave!COP_CTA_DIF_TC_CONTRA)
 c_redg.Text = Trim(cop_llave!cop_cta_red_FAVOR)
 c_redp.Text = Trim(cop_llave!cop_cta_red_contra)
 'txtCampo1.Text = Format(LK_FECHA_COP1, "dd/mm/yyyy")
 txtCampo1.Text = Format(cop_llave!cop_fecha_proceso, "dd/mm/yyyy")
 txtCampo1.Mask = "##/##/####"
 'txtCampo2.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
 act_mes.Caption = Format(cop_llave!cop_nro_mes, "00")
 act_ano.Caption = Format(cop_llave!cop_fecha_proceso, "yyyy")
 txtCampo2.Text = Format(cop_llave!COP_FECHA_PROCESO2, "dd/mm/yyyy")
 txtCampo2.Mask = "##/##/####"
 txtnivelmax.Text = fila - 1
 If cop_llave!COP_FLAG_MAYORIZACION = "A" Then
'   lblestado.Caption = "Contabilidad Mayorizada"
 Else
'   lblestado.Caption = "En Proceso. . ."
 End If
 If Not ver_cuentas.EOF Then
  For fila = 1 To 6
    txtCopar(fila - 1).Enabled = False
  Next fila
  txtnivelmax.Enabled = False
End If
End If
If cop_llave!COP_FLAG_PASE = "A" Then
 Che_pase.FillColor = QBColor(10)
End If
If cop_llave!cop_FLAG_REGC = "A" Then
 Che_regc.FillColor = QBColor(10)
End If
If cop_llave!cop_FLAG_REGV = "A" Then
 che_regv.FillColor = QBColor(10)
End If
If cop_llave!cop_FLAG_CAJA = "A" Then
 Che_caja.FillColor = QBColor(10)
End If
If cop_llave!cop_FLAG_DES = "A" Then
 Che_destinos.FillColor = QBColor(10)
End If
If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 Che_mayor.FillColor = QBColor(10)
End If

For fila = 0 To 12
 MES.ListIndex = fila
 If Val(Left(MES.Text, 2)) = LK_NRO_MES Then
   Exit For
 End If
Next fila
'**********************
tmp = Mid(cop_llave!cop_flag_cc, 1, 1)
chkcc(0).Value = IIf(tmp = 1, 1, 0)
tmp = Mid(cop_llave!cop_flag_cc, 2, 1)
txtdgcc(0).Text = IIf(tmp = 0, 1, tmp)
tmp = Mid(cop_llave!cop_flag_cc, 3, 1)
chkcc(1).Value = IIf(tmp = 1, 1, 0)
tmp = Mid(cop_llave!cop_flag_cc, 4, 1)
txtdgcc(1).Text = IIf(tmp = 0, 1, tmp)
tmp = Mid(cop_llave!cop_flag_cc, 5, 1)
chkcc(2).Value = IIf(tmp = 1, 1, 0)
tmp = Mid(cop_llave!cop_flag_cc, 6, 1)
txtdgcc(2).Text = IIf(tmp = 0, 1, tmp)

pub_cadena = "SELECT DISTINCT cc_tipo FROM centroc WHERE cc_tipo = ?"
Set PSCC = CN.CreateQuery("", pub_cadena)
PSCC(0) = 1
Set RS_CC = PSCC.OpenResultset(rdOpenKeyset, rdConcurValues)
RS_CC.Requery
If Not RS_CC.EOF Then
    chkcc(0).Enabled = False
    txtdgcc(0).Enabled = False
End If
PSCC(0) = 2
RS_CC.Requery
If Not RS_CC.EOF Then
    chkcc(1).Enabled = False
    txtdgcc(1).Enabled = False
End If
PSCC(0) = 3
RS_CC.Requery
If Not RS_CC.EOF Then
    chkcc(2).Enabled = False
    txtdgcc(2).Enabled = False
End If
'**********************


fila = 0
For fila = 0 To ano.ListCount - 1
  ano.ListIndex = fila
  If Val(ano.Text) = Val(Format(LK_FECHA_COP1, "yyyy")) Then
    GoTo fin
  End If
Next fila


fin:
If LK_CODUSU = "ADMIN" Then
  txtCampo1.Visible = True
  txtCampo2.Visible = True
End If
chequea LK_NRO_MES, LK_FECHA_COP1
End Sub

Private Sub MES_Click()
Dim WDATES As Date
If Trim(MES.Text) = "" Then Exit Sub
If Trim(ano.Text) = "" Then Exit Sub
If Left(MES, 2) = "00" Then
 WDATES = CDate("01/01/" & ano.Text)
Else
 WDATES = CDate("01/" & Left(MES, 2) & "/" & ano.Text)
End If
 chequea Val(Left(MES, 2)), WDATES
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  txtCampo2.SetFocus
End If
End Sub

Private Sub txtcampo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdgrabar.SetFocus
End If

End Sub

Private Sub txtnivelmax_Change()
If Val(txtnivelmax) <= 0 Then Exit Sub
For fila = 1 To 6
 If fila <= Val(txtnivelmax) Then
   txtCopar(fila - 1).Enabled = True
 Else
  txtCopar(fila - 1).Enabled = False
  txtCopar(fila - 1).Text = ""
 End If
Next fila
End Sub

Private Sub txtnivelmax_KeyPress(KeyAscii As Integer)
Dim car
If KeyAscii = 8 Then
 Exit Sub
End If
car = Chr(KeyAscii)
If Val(car) < "1" Or Val(car) > "6" Then
  KeyAscii = 0
End If
End Sub
Public Function SON_FECHAS(wf1 As MaskEdBox, wf2 As MaskEdBox) As Boolean
SON_FECHAS = True
If Right(wf1.Text, 2) = "__" Then
  REP_FECHA1 = Left(wf1.Text, 8)
Else
  REP_FECHA1 = Trim(wf1.Text)
End If
If Not IsDate(REP_FECHA1) Then
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 wf1, wf1
    GoTo fin
End If
If Right(wf2.Text, 2) = "__" Then
  REP_FECHA2 = Left(wf2.Text, 8)
Else
  REP_FECHA2 = Trim(wf2.Text)
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 Azul2 wf2, wf2
 GoTo fin
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fechas Invalidadas ..", 48, Pub_Titulo
 Azul2 wf1, wf1
 GoTo fin
End If

Exit Function
fin:
SON_FECHAS = False

End Function


Public Sub LLENA_FECHAS()
  MES.Clear
  MES.AddItem "00" & " - APERTURA"
  MES.AddItem "01" & " - Enero"
  MES.AddItem "02" & " - Febrero"
  MES.AddItem "03" & " - Marzo"
  MES.AddItem "04" & " - Abril"
  MES.AddItem "05" & " - Mayo"
  MES.AddItem "06" & " - Junio"
  MES.AddItem "07" & " - Julio"
  MES.AddItem "08" & " - Agosto"
  MES.AddItem "09" & " - Septiembre"
  MES.AddItem "10" & " - Octubre"
  MES.AddItem "11" & " - Noviembre"
  MES.AddItem "12" & " - Diciembre"
  MES.AddItem "13" & " - AJUSTE 1"
  MES.AddItem "14" & " - AJUSTE 2"
  MES.AddItem "15" & " - CIERRE ANUAL"
  ano.ListIndex = -1
 
End Sub


Public Sub chequea(W_NRO_MES As Integer, W_FECHA_COP1 As Date)
Dim wvalor As String * 1
Dim wcadena As String * 13
OTT:
cmdest.Enabled = True
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  tab_mayor.AddNew
  tab_mayor!TAB_CODCIA = LK_CODCIA
  tab_mayor!TAB_TIPREG = 155
  tab_mayor!TAB_NUMTAB = Format(W_FECHA_COP1, "yyyy")
  tab_mayor!tab_nomlargo = "0000000000000"
  tab_mayor!tab_nomcorto = ""
  tab_mayor!TAB_CONTABLE2 = 0
  tab_mayor!TAB_CODART = 0
  tab_mayor!TAB_CODCLIE = 0
  tab_mayor.Update
  GoTo OTT
  cmdest.Enabled = False
'  MsgBox "Crear tab_tipreg = 155 para seguridad en Meses ", 48, Pub_Titulo
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(W_FECHA_COP1, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, W_NRO_MES + 1, 1)
    If wvalor = "1" Then
       estado.Caption = "CERRADO"
       cmdest.Caption = "Abrir Mes"
    Else
       estado.Caption = "DISPONIBLE"
       cmdest.Caption = "Cerrar Mes"
    End If
End If

End Sub

Private Sub txtpas_KeyPress(KeyAscii As Integer)
Dim WPASS As String

WPASS = PUB_CLAVE
If KeyAscii = 27 Then
  FRACLAVE.Visible = False
  Exit Sub
End If
If KeyAscii = 13 Then
  If Left(cmdest.Caption, 1) <> "A" Then
    OEMES Left(MES.Text, 2), "01/01/" & Trim(ano.Text), 1
     MsgBox "Mes Cerrado!!!", 48, Pub_Titulo
     txtpas.Text = ""
     FRACLAVE.Visible = False
     cmdest.SetFocus
  Else
    If WPASS <> Trim(txtpas.Text) Then
         MsgBox "Clave de Acceso Incorrecta.........", 48, Pub_Titulo
         txtpas.Text = ""
          FRACLAVE.Visible = False
         cmdest.SetFocus
         Exit Sub
    End If
    OEMES Left(MES.Text, 2), "01/01/" & Trim(ano.Text), 0
    MsgBox "**** Mes esta Disponible ****", 48, Pub_Titulo
    FRACLAVE.Visible = False
    txtpas.Text = ""
    cmdest.SetFocus
  End If
End If
End Sub

Public Sub OEMES(W_NRO_MES As Integer, W_FECHA_COP1 As Date, OE As Integer)
Dim WTEMP_VAR As String * 13
Dim wvalor As String * 1
Dim wcadena As String * 13
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
  Exit Sub
End If
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(W_FECHA_COP1, "yyyy")) Then
         WTEMP_VAR = Trim(tab_mayor!tab_nomlargo)
   GoTo ACT
       End If
      tab_mayor.MoveNext
    Loop
    MsgBox "NO procede , reinicie la maquina . crear tab_tipreg=155 "
Exit Sub
ACT:
Dim WTEMP0 As String * 1
Dim WTEMP1 As String * 1
Dim WTEMP2 As String * 1
Dim WTEMP3 As String * 1
Dim WTEMP4 As String * 1
Dim WTEMP5 As String * 1
Dim WTEMP6 As String * 1
Dim WTEMP7 As String * 1
Dim WTEMP8 As String * 1
Dim WTEMP9 As String * 1
Dim WTEMP10 As String * 1
Dim WTEMP11 As String * 1
Dim WTEMP12 As String * 1


WTEMP0 = Mid(WTEMP_VAR, 1, 1)
WTEMP1 = Mid(WTEMP_VAR, 2, 1)
WTEMP2 = Mid(WTEMP_VAR, 3, 1)
WTEMP3 = Mid(WTEMP_VAR, 4, 1)
WTEMP4 = Mid(WTEMP_VAR, 5, 1)
WTEMP5 = Mid(WTEMP_VAR, 6, 1)
WTEMP6 = Mid(WTEMP_VAR, 7, 1)
WTEMP7 = Mid(WTEMP_VAR, 8, 1)
WTEMP8 = Mid(WTEMP_VAR, 9, 1)
WTEMP9 = Mid(WTEMP_VAR, 10, 1)
WTEMP10 = Mid(WTEMP_VAR, 11, 1)
WTEMP11 = Mid(WTEMP_VAR, 12, 1)
WTEMP12 = Mid(WTEMP_VAR, 13, 1)
If W_NRO_MES = 0 Then WTEMP0 = OE
If W_NRO_MES = 1 Then WTEMP1 = OE
If W_NRO_MES = 2 Then WTEMP2 = OE
If W_NRO_MES = 3 Then WTEMP3 = OE
If W_NRO_MES = 4 Then WTEMP4 = OE
If W_NRO_MES = 5 Then WTEMP5 = OE
If W_NRO_MES = 6 Then WTEMP6 = OE
If W_NRO_MES = 7 Then WTEMP7 = OE
If W_NRO_MES = 8 Then WTEMP8 = OE
If W_NRO_MES = 9 Then WTEMP9 = OE
If W_NRO_MES = 10 Then WTEMP10 = OE
If W_NRO_MES = 11 Then WTEMP11 = OE
If W_NRO_MES = 12 Then WTEMP12 = OE
WTEMP_VAR = WTEMP0 & WTEMP1 & WTEMP2 & WTEMP3 & WTEMP4 & WTEMP5 & WTEMP6 & WTEMP7 & WTEMP8 & WTEMP9 & WTEMP10 & WTEMP11 + WTEMP12
tab_mayor.Edit
tab_mayor!tab_nomlargo = WTEMP_VAR
tab_mayor.Update
If OE = "1" Then
   estado.Caption = "CERRADO"
   cmdest.Caption = "Abrir Mes"
Else
   estado.Caption = "DISPONIBLE"
   cmdest.Caption = "Cerrar Mes"
End If


End Sub

Private Sub txtpas_LostFocus()
  FRACLAVE.Visible = False
End Sub

Private Sub chkcc_Click(Index As Integer)
    If chkcc(2).Value = 1 And Index = 2 Then
        If chkcc(1).Value = 0 Then
            MsgBox "Tiene que marcar el Chek 2"
            chkcc(2).Value = 0
            Exit Sub
        End If
    End If
    If chkcc(Index).Value = 1 Then
        txtdgcc(Index).Enabled = True
        If txtdgcc(Index).Visible Then txtdgcc(Index).SetFocus
        If Index = 1 Then chkcc(2).Enabled = True
    Else
        txtdgcc(Index).Enabled = False
        txtdgcc(Index).Text = "1"
        If Index = 1 Then chkcc(2).Enabled = False: txtdgcc(2).Enabled = False: txtdgcc(2).Text = "1": chkcc(2).Value = 0
    End If
End Sub

Private Sub txtdgcc_LostFocus(Index As Integer)
    If txtdgcc(Index).Text = "" And chkcc(Index).Value = 1 Then
        MsgBox "Ingrese el Numero de Digitos del CC"
        txtdgcc(Index).SetFocus
        Exit Sub
    End If
    If Val(txtdgcc(Index).Text) > 10 Or Val(txtdgcc(Index).Text) < 1 Then
        MsgBox "Ingrese solo digitos. Sólo (1 - 9)"
        txtdgcc(Index).SetFocus
        txtdgcc(Index).Text = "1"
    End If
End Sub
Private Sub txtdgcc_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii < 49 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        MsgBox "Ingrese solo digitos. Sólo (1 - 9)"
        txtdgcc(Index).SetFocus
        txtdgcc(Index).Text = ""
        KeyAscii = 49
    End If
    If txtdgcc(Index).Text = "" And KeyAscii = 13 Then
        MsgBox "Ingrese solo digitos. Sólo (1 - 9)"
        txtdgcc(Index).SetFocus
    End If
End Sub
