VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTareas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Programación de Tareas"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   410
      Left            =   1335
      TabIndex        =   19
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   410
      Left            =   2490
      TabIndex        =   18
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   410
      Left            =   3630
      TabIndex        =   17
      Top             =   3825
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   0
      TabIndex        =   6
      Top             =   -15
      Width           =   4740
      Begin MSComCtl2.DTPicker dtpHora 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "hh:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2880
         TabIndex        =   15
         Tag             =   "Ingrese la hora exacta de la tarea"
         Top             =   1455
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   -2147483639
         CustomFormat    =   "hh:mm:ss AMPM"
         Format          =   61800450
         UpDown          =   -1  'True
         CurrentDate     =   37168
      End
      Begin VB.TextBox txtCodigo 
         Height          =   285
         Left            =   930
         TabIndex        =   0
         Tag             =   "Ingrese el codigo de la tarea/<F1> por Ayuda"
         Top             =   225
         Width           =   675
      End
      Begin VB.TextBox txtDesc 
         Height          =   285
         Left            =   330
         MaxLength       =   20
         TabIndex        =   1
         Tag             =   "Ingrese una descripción corta de la tarea"
         Top             =   885
         Width           =   4065
      End
      Begin VB.TextBox txtObs 
         Height          =   750
         Left            =   270
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Tag             =   "Ingrese la descrpción detallada de la tarea."
         Top             =   2160
         Width           =   4140
      End
      Begin VB.TextBox txtDias 
         Height          =   285
         Left            =   300
         MaxLength       =   2
         TabIndex        =   2
         Tag             =   "Ingrese los dias de anticipación a la fecha establecida"
         Text            =   "1"
         Top             =   1485
         Width           =   855
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   300
         Left            =   1140
         Max             =   1
         Min             =   10
         TabIndex        =   7
         Top             =   1455
         Value           =   1
         Width           =   225
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Tarea Común"
         Height          =   195
         Left            =   330
         TabIndex        =   5
         Top             =   3060
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   315
         Left            =   1470
         TabIndex        =   3
         Tag             =   "Ingrese la fecha a realizarse la tarea"
         Top             =   1455
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   -2147483635
         CalendarTitleForeColor=   -2147483639
         Format          =   61800449
         CurrentDate     =   37162
      End
      Begin VB.Label Label3 
         Caption         =   "Hora:"
         Height          =   180
         Left            =   2910
         TabIndex        =   14
         Top             =   1245
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Codigo"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   210
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Left            =   1485
         TabIndex        =   12
         Top             =   1245
         Width           =   450
      End
      Begin VB.Label lt 
         AutoSize        =   -1  'True
         Caption         =   "Descripción  Corta:"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   11
         Top             =   645
         Width           =   1350
      End
      Begin VB.Label lt 
         AutoSize        =   -1  'True
         Caption         =   "Observación:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   10
         Top             =   1860
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Diaz de Plazo :"
         Height          =   195
         Left            =   285
         TabIndex        =   9
         Top             =   1245
         Width           =   1065
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha :"
         Height          =   225
         Left            =   2160
         TabIndex        =   8
         Top             =   885
         Width           =   1170
      End
   End
   Begin VB.Label lblMensage 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   15
      TabIndex        =   16
      Top             =   3465
      Width           =   4710
   End
End
Attribute VB_Name = "frmTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NeworEdit As Integer
Private Sub cmdAceptar_Click()
 
If SaveData() Then
 txtDesc = ""
 txtCodigo = ""
 txtDias = 1
 dtpFecha = Date
 dtpHora = Time
 txtObs = ""
 cmdAceptar.Default = False
 txtCodigo.SetFocus
 NeworEdit = 0
End If
End Sub
Private Function SaveData() As Boolean
Dim sql As String
If NeworEdit = 1 Then
 sql = "INSERT INTO TAREAS (IDTAREA,USER_TAREA,DESCTAREA,HORA,FECHA, " & _
       "DIAS,OBS) VALUES ('" & Trim(txtCodigo) & "','" & LK_CODUSU & "','" & txtDesc & "','" & dtpHora & "','" & _
       Format(dtpFecha, "dd/mm/yyyy") & "','" & txtDias & "','" & txtObs & "')"
ElseIf NeworEdit = 2 Then
 sql = "UPDATE TAREAS SET DESCTAREA='" & txtDesc & "', " & _
 "HORA='" & TimeValue(dtpHora) & "',FECHA='" & Format(dtpFecha, "dd/mm/yyyy") & "',DIAS='" & _
 txtDias & "',OBS='" & txtObs & "' WHERE idtarea='" & Trim(txtCodigo) & "'"
End If
 On Error GoTo ErrorSave
 cnT.Execute sql
 SaveData = True
 NeworEdit = 0
 Exit Function
ErrorSave:
 MsgBox Err.Description
 SaveData = False
End Function
Private Sub cmdcancelar_Click()
If Not txtCodigo = "" Then
  cmdAceptar.Default = False
  txtDesc = ""
  txtCodigo = ""
  txtDias = 1
  dtpFecha = Format(Date, "dd/mm/yyyy")
  txtObs = ""
  txtCodigo.SetFocus
Else
Unload Me
End If
End Sub
Private Sub cmdEliminar_Click()
Dim sql As String
sql = "DELETE FROM TAREAS WHERE IDTAREA='" & Trim(txtCodigo) & "'"
 On Error GoTo ErrorSave
 CN.Execute sql
 txtDesc = ""
 txtCodigo = ""
 txtDias = 1
 dtpHora = Time
 dtpFecha = Format(Date, "dd/mm/yyyy")
 txtObs = ""
 txtCodigo.SetFocus
 Exit Sub
ErrorSave:
 MsgBox Err.Description
End Sub
Private Sub dtpFecha_GotFocus()
lblMensage = dtpFecha.Tag
End Sub
Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 If CInt(DateDiff("d", Date, dtpFecha)) >= Val(txtDias) Then
  dtpHora.SetFocus
 Else
  MsgBox "El numero de dias ingresados no es " & _
  "valido para la fecha ingresada, cambie " & _
  "la fecha o el numero de dias", _
         vbInformation, "Validadción de Fecha"
  dtpFecha.SetFocus
 End If
End If
End Sub
Private Sub dtpHora_GotFocus()
lblMensage = dtpHora.Tag
End Sub
Private Sub dtpHora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  txtObs.SetFocus
End If
End Sub
Private Sub Form_Load()
Dim str As String
Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=BDATOS;Origen de datos=server"
'str = "Provider=SQLOLEDB;Data Source=SERVER;initial catalog=BDATOS;user id=sa;password=;"
cnT.Open str
dtpHora = Time
dtpFecha = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set rsT = Nothing
Set cnT = Nothing
End Sub
Private Sub txtCodigo_GotFocus()
lblMensage = txtCodigo.Tag
End Sub
Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
 msg = "HELP"
 frmListTareas.Show 1, Me
End If
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 ShowData
End If
End Sub
Sub ShowData()
Dim sql As String
sql = "select * from tareas where idtarea='" & Trim((txtCodigo)) & "'"
 If rsT.State = adStateOpen Then
  Set rsT = Nothing
 End If
 rsT.CursorLocation = adUseClient
 rsT.Open sql, cnT, adOpenStatic, adLockOptimistic, adCmdText
 If Not rsT.RecordCount = 0 Then
 NeworEdit = 2
 txtDesc = rsT(2)
 txtDias = rsT(5)
 dtpFecha = rsT(3)
 dtpHora = rsT(4)
 txtObs = rsT(6)
 Set rsT = Nothing
 txtDesc.SetFocus
 Else
 NeworEdit = 1
 Set rsT = Nothing
 txtDesc.SetFocus
 End If
End Sub
Private Sub txtDesc_GotFocus()
lblMensage = txtDesc.Tag
End Sub
Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 KeyAscii = 0
 txtDias.SetFocus
End If
End Sub
Private Sub txtDias_GotFocus()
lblMensage = txtDias.Tag
End Sub
Private Sub txtDias_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
 KeyAscii = 0
 dtpFecha.SetFocus
End If
End Sub
Private Sub txtObs_GotFocus()
cmdAceptar.Default = True
lblMensage = txtObs.Tag
End Sub
Private Sub VScroll1_Change()
txtDias = VScroll1.Value
End Sub
