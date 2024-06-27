VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmDepura 
   BackColor       =   &H00000000&
   Caption         =   "Depuración "
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "FrmDepura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmddepura 
      Caption         =   "Depuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Picture         =   "FrmDepura.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin MSMask.MaskEdBox txtCampo2 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   0
      ForeColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Fechas :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   750
   End
   Begin VB.Label ESTADO 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label lblcompa 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Estado: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Compañia :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "OSBusiness"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Depuración de Registros Extornados en OSBusinnes - SQLServer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmDepura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PSDEPU_LLAVE2 As rdoQuery
Dim depu_llave2 As rdoResultset
Dim PSDEPU_LLAVE1 As rdoQuery
Dim depu_llave1 As rdoResultset
Dim PSDEPU_LLAVE3 As rdoQuery
Dim depu_llave3 As rdoResultset

  
Private Sub cmddepura_Click()
' VARIABLE PARA COMPAñIA DE RESPALDO
Dim RESP_CODCIA As String * 2
RESP_CODCIA = ""
RESP_CODCIA = "D" & Trim(Str(Val(LK_CODCIA)))
'========================
' MANTEMIMIENTO DE FACART
'========================
'** TODO MENOS VENTAS
pub_cadena = "SELECT FAR_TIPMOV, FAR_ESTADO, FAR_CODCIA FROM FACART WHERE FAR_CODCIA = ? AND (FAR_FECHA >= ? AND FAR_FECHA <= ?) AND FAR_ESTADO = 'E'  AND FAR_TIPMOV <> 10 "
Set PSDEPU_LLAVE1 = CN.CreateQuery("", pub_cadena)
PSDEPU_LLAVE1(0) = 0
PSDEPU_LLAVE1(1) = LK_FECHA_DIA
PSDEPU_LLAVE1(2) = LK_FECHA_DIA
Set depu_llave1 = PSDEPU_LLAVE1.OpenResultset(rdOpenKeyset, rdConcurValues)
PSDEPU_LLAVE1(0) = LK_CODCIA
PSDEPU_LLAVE1(1) = txtCampo1.text
PSDEPU_LLAVE1(2) = txtCampo2.text
depu_llave1.Requery
Do Until depu_llave1.EOF
 DoEvents
 If Val(depu_llave1.AbsolutePosition) > Val(depu_llave1.RowCount) Then
   Exit Do
 End If
 ESTADO.Caption = depu_llave1.AbsolutePosition & "/" & depu_llave1.RowCount
 If depu_llave1!far_ESTADO <> "E" Then
  MsgBox "NO PROCEDE..ESTE REGISTRO"
  GoTo PASA
 End If
 If depu_llave1!far_TIPMOV = 10 Then
   MsgBox "NO PROCEDE..ESTE REGISTRO"
   GoTo PASA
  End If
 depu_llave1.Edit
 depu_llave1!FAR_CODCIA = RESP_CODCIA
 depu_llave1.Update
PASA:
 depu_llave1.MoveNext
Loop

'** VENTAS
pub_cadena = "SELECT FAR_ESTADO, FAR_CODCIA FROM FACART WHERE FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_ESTADO = 'E' "
Set PSDEPU_LLAVE3 = CN.CreateQuery("", pub_cadena)
PSDEPU_LLAVE3(0) = 0
PSDEPU_LLAVE3(1) = 0
PSDEPU_LLAVE3(2) = 0
PSDEPU_LLAVE3(3) = 0
Set depu_llave3 = PSDEPU_LLAVE3.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_CODCIA, FAR_ESTADO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_ESTADO = 'N' "
Set PSDEPU_LLAVE2 = CN.CreateQuery("", pub_cadena)
PSDEPU_LLAVE2(0) = 0
PSDEPU_LLAVE2(1) = 0
PSDEPU_LLAVE2(2) = 0
PSDEPU_LLAVE2(3) = 0
Set depu_llave2 = PSDEPU_LLAVE2.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT DISTINCT FAR_FBG, FAR_NUMSER, FAR_NUMFAC FROM FACART WHERE FAR_CODCIA = ? AND (FAR_FECHA >= ? AND FAR_FECHA <= ?) AND FAR_ESTADO = 'E'  AND FAR_TIPMOV = 10 GROUP BY FAR_CODCIA, FAR_FBG, FAR_NUMSER, FAR_NUMFAC"
Set PSDEPU_LLAVE1 = CN.CreateQuery("", pub_cadena)
PSDEPU_LLAVE1(0) = 0
PSDEPU_LLAVE1(1) = LK_FECHA_DIA
PSDEPU_LLAVE1(2) = LK_FECHA_DIA
Set depu_llave1 = PSDEPU_LLAVE1.OpenResultset(rdOpenKeyset, rdConcurValues)
PSDEPU_LLAVE1(0) = LK_CODCIA
PSDEPU_LLAVE1(1) = txtCampo1.text
PSDEPU_LLAVE1(2) = txtCampo2.text
depu_llave1.Requery
fila = 0
Do Until depu_llave1.EOF
 DoEvents
 If Val(depu_llave1.AbsolutePosition) > Val(depu_llave1.RowCount) Then
   Exit Do
 End If
' Stop
 ESTADO.Caption = depu_llave1.AbsolutePosition & "/" & depu_llave1.RowCount
' If depu_llave1!far_ESTADO <> "E" Then
'  MsgBox "NO PROCEDE..ESTE REGISTRO"
'  GoTo PASA2
' End If
 ' VERIFICAR SI NO ESTA ANULADO
 PSDEPU_LLAVE2(0) = LK_CODCIA
 PSDEPU_LLAVE2(1) = depu_llave1!far_FBG
 PSDEPU_LLAVE2(2) = depu_llave1!far_NUMSER
 PSDEPU_LLAVE2(3) = depu_llave1!far_NUMFAC
 depu_llave2.Requery
 If depu_llave2.EOF Then GoTo PASA2
 
 ' ACTUALIZA O PASE A OTRA CIA
 PSDEPU_LLAVE3(0) = LK_CODCIA
 PSDEPU_LLAVE3(1) = depu_llave1!far_FBG
 PSDEPU_LLAVE3(2) = depu_llave1!far_NUMSER
 PSDEPU_LLAVE3(3) = depu_llave1!far_NUMFAC
 depu_llave3.Requery
 Do Until depu_llave3.EOF
   If Val(depu_llave3.AbsolutePosition) > Val(depu_llave3.RowCount) Then
     Exit Do
   End If
   If depu_llave3!far_ESTADO <> "E" Then
    MsgBox "NO PROCEDE..ESTE REGISTRO"
    GoTo PASA2
   End If
   depu_llave3.Edit
   depu_llave3!FAR_CODCIA = RESP_CODCIA
  ' RESP_CODCIA = "D3"
   depu_llave3.Update
   fila = fila + 1
   
   depu_llave3.MoveNext
 Loop
 

  
 
PASA2:
 depu_llave1.MoveNext
Loop



MsgBox "PROCESO TERMINADO - total " & fila, 48, Pub_Titulo
 


End Sub

Private Sub Form_Load()
lblcompa.Caption = Trim(MDIForm1.txtcia.Caption)
txtCampo1.Mask = "##/##/####"
txtCampo1.text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtCampo2.Mask = "##/##/####"
txtCampo2.text = Format(LK_FECHA_DIA, "dd/mm/yyyy")



End Sub
