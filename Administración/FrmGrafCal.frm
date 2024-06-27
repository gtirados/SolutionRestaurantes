VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form FrmGrafCal 
   Caption         =   "Visor de Calendario"
   ClientHeight    =   6510
   ClientLeft      =   135
   ClientTop       =   945
   ClientWidth     =   11625
   Icon            =   "FrmGrafCal.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmGrafCal.frx":0442
   ScaleHeight     =   6510
   ScaleWidth      =   11625
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   7800
      TabIndex        =   45
      Top             =   480
      Width           =   3735
      Begin VB.TextBox txtobs 
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Text            =   "Text2"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtdescrip 
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lt 
         Caption         =   "Observación:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lt 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame f1 
      Caption         =   "Uso de Calendario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
      Begin VB.TextBox txtayuda 
         Height          =   3375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   40
         Text            =   "FrmGrafCal.frx":247C
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   1320
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdayuda 
      Caption         =   "&Ayuda"
      Height          =   375
      Left            =   6480
      TabIndex        =   32
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton CmdRetornar 
      Caption         =   "&Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      Picture         =   "FrmGrafCal.frx":261F
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5880
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   16776960
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      FocusRect       =   2
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   2
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   3
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   4
      Left            =   5760
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   5
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   6
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   7
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   8
      Left            =   5760
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   11
      Left            =   3840
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   4320
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   10
      Left            =   1920
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid GridEne 
      Height          =   1455
      Index           =   12
      Left            =   5760
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   2566
      _Version        =   393216
      ForeColor       =   0
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorBkg    =   12632256
      Redraw          =   -1  'True
      AllowBigSelection=   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      MergeCells      =   1
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Programacion de Tareas"
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
      Height          =   255
      Left            =   7800
      TabIndex        =   44
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disponible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   4920
      TabIndex        =   43
      Top             =   5955
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   42
      Top             =   5955
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "Leyenda :"
      Height          =   255
      Left            =   0
      TabIndex        =   41
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   39
      Top             =   5955
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   38
      Top             =   5955
      Width           =   135
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   37
      Top             =   5955
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fec.Procesada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3480
      TabIndex        =   36
      Top             =   5955
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fec. Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   35
      Top             =   5955
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Feriado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1320
      TabIndex        =   34
      Top             =   5955
      Width           =   780
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   33
      Top             =   5880
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Calendario del Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   2040
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Noviembre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   11
      Left            =   4320
      TabIndex        =   23
      Top             =   4080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Diciembre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   12
      Left            =   6120
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Octubre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   10
      Left            =   2400
      TabIndex        =   21
      Top             =   4080
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Setiembre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   480
      TabIndex        =   20
      Top             =   4080
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Julio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   4560
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Agosto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   6360
      TabIndex        =   18
      Top             =   2400
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Junio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   2520
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   600
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Marzo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   4440
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Abril"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   6480
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Febrero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   2520
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label LblMes 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Enero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   600
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Linea1 
      BackColor       =   &H008B4914&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label Linea1 
      BackColor       =   &H008B4914&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   25
      Top             =   2400
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label Linea1 
      BackColor       =   &H008B4914&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.Label LblAno 
      AutoSize        =   -1  'True
      Caption         =   "           "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   28
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "FrmGrafCal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LOC_MES As Integer
Dim temporal

Private Sub cmdayuda_Click()
F1.Left = 7680
F1.Top = 1440
F1.Height = 3735
F1.Width = 1815
F1.Visible = True
txtayuda.SetFocus

End Sub

Private Sub CmdRetornar_Click()
Unload FrmGrafCal
GENCAL.Show 1
End Sub


Public Sub GEN_MES(GRIDMES As MSFlexGrid, WDiaInc As Integer, WTotDia As Integer)
Dim diafin As Integer
Dim swa, swa2 As Integer
Dim Wrow As Integer
Dim cont As Integer
Dim WCOL As Integer
Dim NumDias, i As Integer
GRIDMES.Clear
GRIDMES.Rows = 6
GRIDMES.Cols = 8
GRIDMES.Col = 1
GRIDMES.Row = 0
GRIDMES.text = "D"
GRIDMES.Col = 2
GRIDMES.text = "L"
GRIDMES.Col = 3
GRIDMES.text = "M"
GRIDMES.Col = 4
GRIDMES.text = "M"
GRIDMES.Col = 5
GRIDMES.text = "J"
GRIDMES.Col = 6
GRIDMES.text = "V"
GRIDMES.Col = 7
GRIDMES.text = "S"
swa = 250 ' 280
swa2 = 230 ' 280
GRIDMES.Width = 1900
GRIDMES.Height = 1450
GRIDMES.ColWidth(0) = 1
GRIDMES.ColWidth(1) = swa
GRIDMES.ColWidth(2) = swa
GRIDMES.ColWidth(3) = swa
GRIDMES.ColWidth(4) = swa
GRIDMES.ColWidth(5) = swa
GRIDMES.ColWidth(6) = swa
GRIDMES.ColWidth(7) = swa

GRIDMES.RowHeight(1) = swa2
GRIDMES.RowHeight(2) = swa2
GRIDMES.RowHeight(3) = swa2
GRIDMES.RowHeight(4) = swa2
GRIDMES.RowHeight(5) = swa2


'WDiaInc = 1
'WTotDia = 31
cont = 0
WCOL = WDiaInc
NumDias = 1
Wrow = 1
Do Until NumDias = WTotDia + 1
    If WCOL = 8 Then
       Wrow = Wrow + 1
       WCOL = 1
    End If
    If Wrow = 6 Then
       Wrow = 1
    End If
    GRIDMES.Row = Wrow
    GRIDMES.Col = WCOL
    GRIDMES.text = NumDias
    diafin = NumDias
    NumDias = NumDias + 1
    WCOL = WCOL + 1
Loop
GRIDMES.Tag = "01" & Format(diafin, "00")
End Sub
Public Sub MARCA_FER(GRIDMES As MSFlexGrid, WDiaFer)
Dim WCOL As Integer
Dim WFIL As Integer
WCOL = 1
WFIL = 0
For WCOL = 1 To 7
  For WFIL = 1 To 5
    GRIDMES.Row = WFIL
    GRIDMES.Col = WCOL
    If Trim(GRIDMES.text) = Trim(WDiaFer) Then
        GRIDMES.CellForeColor = vbWhite
        GRIDMES.CellBackColor = vbRed
       Exit Sub
    End If
    Next WFIL
Next WCOL
End Sub
Public Sub MARCA_CERRADO(GRIDMES As MSFlexGrid, WDiaFer)
Dim WCOL As Integer
Dim WFIL As Integer
WCOL = 1
WFIL = 0
For WCOL = 1 To 7
  For WFIL = 1 To 5
    GRIDMES.Row = WFIL
    GRIDMES.Col = WCOL
    If Trim(GRIDMES.text) = Trim(WDiaFer) Then
        GRIDMES.CellForeColor = vbBlack
        GRIDMES.CellBackColor = vb3DLight
       Exit Sub
    End If
    Next WFIL
Next WCOL
End Sub

Public Sub MARCA_INICIO(GRIDMES As MSFlexGrid, WDiaFer)
Dim WCOL As Integer
Dim WFIL As Integer
WCOL = 1
WFIL = 0
For WCOL = 1 To 7
  For WFIL = 1 To 5
    GRIDMES.Row = WFIL
    GRIDMES.Col = WCOL
    If Trim(GRIDMES.text) = Trim(WDiaFer) Then
        GRIDMES.CellBackColor = vbCyan
        GRIDMES.CellForeColor = vbBlack
       Exit Sub
    End If
    Next WFIL
Next WCOL
End Sub


Private Sub Form_Load()
Dim i As Integer
Dim cade As String
Dim MESi(12) As Integer
Dim MESf(12) As Integer
Dim WMES As Integer
Dim TEMMES As Integer
LOC_MES = 0
LblAno.Caption = PUB_CAL_ANO
PUB_CAL_INI = "1/1/" & PUB_CAL_ANO
PUB_CAL_FIN = "12/31/" & PUB_CAL_ANO
pu_codcia = LK_CODCIA
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_CAL_LLAVE 1

i = 1
TEMMES = 0
cal_llave.MoveFirst
WMES = Month(cal_llave!CAL_FECHA)
MESi(i) = Weekday(cal_llave!CAL_FECHA)
Do Until cal_llave.EOF
   TEMMES = TEMMES + 1
   If WMES <> Month(cal_llave!CAL_FECHA) Then
        WMES = Month(cal_llave!CAL_FECHA)
        i = i + 1
        MESi(i) = Weekday(cal_llave!CAL_FECHA)
        MESf(i - 1) = TEMMES - 1
        TEMMES = 1
   End If
   If i = 12 Then
      MESf(12) = 31
      Exit Do
   End If
   
   cal_llave.MoveNext
Loop
For i = 1 To 12
    GEN_MES GridEne(i), MESi(i), MESf(i)
Next i
cal_llave.MoveFirst
Do Until cal_llave.EOF
   If cal_llave!CAL_LABORABLE = "N" Then
       WMES = Month(cal_llave!CAL_FECHA)
       TEMMES = Day(cal_llave!CAL_FECHA)
       MARCA_FER GridEne(WMES), TEMMES
   End If
   If cal_llave!CAL_INDICE = "1" Then
       WMES = Month(cal_llave!CAL_FECHA)
       TEMMES = Day(cal_llave!CAL_FECHA)
       MARCA_INICIO GridEne(WMES), TEMMES
   End If
   If cal_llave!CAL_INDICE = "3" Then
       WMES = Month(cal_llave!CAL_FECHA)
       TEMMES = Day(cal_llave!CAL_FECHA)
       MARCA_CERRADO GridEne(WMES), TEMMES
   End If
   cal_llave.MoveNext
Loop

For i = 1 To 12
    LblMes(i).Visible = True
    GridEne(i).Visible = True
    If i < 4 Then
       Linea1(i - 1).Visible = True
    End If
Next i
'FrmGrafCal.cmdtipocambio.Caption = "Tipo de Cambio US$"
CenterMe FrmGrafCal
Screen.MousePointer = 0
End Sub

Private Sub GridEne_DblClick(Index As Integer)
Dim WSLABOR As String * 1
Dim WSMES, WSDIA, WSANO As String
Dim WSFECHA As String
 If GridEne(Index).text = "" Or GridEne(Index).CellBackColor = vb3DLight Then
    Exit Sub
 End If
 If GridEne(Index).CellBackColor = vbCyan And GridEne(Index).CellForeColor = vbBlack Then
    MsgBox "Desactive fecha de Inicio con < Clik Derecho  > ...", 48, Pub_Titulo
    Exit Sub
 End If
  
 If GridEne(Index).CellForeColor = vbWhite Then
    GridEne(Index).CellForeColor = vbBlack
    GridEne(Index).CellBackColor = vbWhite
    WSLABOR = "S"
 Else
    GridEne(Index).CellForeColor = vbWhite
    GridEne(Index).CellBackColor = vbRed
    WSLABOR = "N"
 End If
 WSDIA = Trim(GridEne(Index).text)
 WSMES = Index
 WSANO = Trim(LblAno.Caption)
 SQ_OPER = 1
 WSFECHA = WSDIA & "/" & WSMES & "/" & WSANO
 PUB_CAL_INI = WSFECHA
 PUB_CAL_FIN = WSFECHA
 PUB_CODCIA = LK_CODCIA
 LEER_CAL_LLAVE 1
 cal_llave.Edit
 cal_llave!CAL_LABORABLE = WSLABOR
 cal_llave.Update
End Sub

Private Sub GridEne_GotFocus(Index As Integer)
'gridigv.Visible = False
End Sub

Private Sub GridEne_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim WSINDICE As String * 1
Dim WSMES, WSDIA, WSANO As String
Dim WSFECHA As String
If Button <> 2 Then
   GoTo FIN
End If
If GridEne(Index).text = "" Or GridEne(Index).CellBackColor = vb3DLight Then
    Exit Sub
End If
If GridEne(Index).CellBackColor = vbRed Then
    MsgBox "Fecha Feriada no Procede  Activela con < DobleClik Derecho  > ...", 48, Pub_Titulo
    Exit Sub
End If
 

 If GridEne(Index).CellForeColor = vbBlack And GridEne(Index).CellBackColor = vbCyan Then
    GridEne(Index).CellForeColor = vbBlack
    GridEne(Index).CellBackColor = vbWhite
    WSINDICE = "0"
 Else
    GridEne(Index).CellForeColor = vbBlack
    GridEne(Index).CellBackColor = vbCyan
    WSINDICE = "1"
 End If
 
 WSDIA = Trim(GridEne(Index).text)
 WSMES = Index
 WSANO = Trim(LblAno.Caption)
 SQ_OPER = 1
 WSFECHA = WSDIA & "/" & WSMES & "/" & WSANO
 PUB_CAL_INI = WSFECHA
 PUB_CAL_FIN = WSFECHA
 PUB_CODCIA = LK_CODCIA
 LEER_CAL_LLAVE 1
 cal_llave.Edit
 cal_llave!CAL_INDICE = WSINDICE
 cal_llave.Update
par_llave.Requery
If WSINDICE = "1" Then
  par_llave.Edit
  par_llave!PAR_FECHA_DIA = WSFECHA
  par_llave.Update
End If

FIN:
End Sub




Private Sub LblMes_Click(Index As Integer)
Dim wm As String
LOC_MES = Index
cmdtipocambio.Caption = ""
wm = ""
If Index = 1 Then
wm = "Enero"
ElseIf Index = 2 Then
wm = "Febrero"
ElseIf Index = 3 Then
wm = "Marzo"
ElseIf Index = 4 Then
wm = "Abril"
ElseIf Index = 5 Then
wm = "Mayo"
ElseIf Index = 6 Then
wm = "Junio"
ElseIf Index = 7 Then
wm = "Julio"
ElseIf Index = 8 Then
wm = "Agosto"
ElseIf Index = 9 Then
wm = "Septiembre"
ElseIf Index = 10 Then
wm = "Octubre"
ElseIf Index = 11 Then
wm = "Nobiembre"
ElseIf Index = 12 Then
wm = "Diciembre"
End If
If wm <> "" Then
  cmdtipocambio.Caption = "Tipo de Cambio US$ del Mes - " & UCase(wm)
  gridigv.Visible = False
Else
  cmdtipocambio.Caption = ""
End If

End Sub


Private Sub TEXTOVAR_GotFocus()
 temporal = gridigv.TextMatrix(gridigv.Row, gridigv.Col)
End Sub


Private Sub txtayuda_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  F1.Visible = False
  cmdayuda.SetFocus
End If
End Sub

Private Sub txtayuda_LostFocus()
  F1.Visible = False
  cmdayuda.SetFocus

End Sub
Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static valor
  Dim car As String
 ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

