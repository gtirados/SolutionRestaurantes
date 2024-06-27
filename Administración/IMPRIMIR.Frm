VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmImprime 
   Caption         =   "IMPRIMIR"
   ClientHeight    =   3120
   ClientLeft      =   945
   ClientTop       =   2025
   ClientWidth     =   6570
   Icon            =   "IMPRIMIR.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   6570
   Begin VB.CommandButton Command3 
      Caption         =   "Actualizar todos los  Articulos con Costo Promedio "
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   480
      Width           =   3855
   End
   Begin Crystal.CrystalReport crystal1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin VB.Frame Frame2 
      Caption         =   "Imprimir  TRANSACCIONES  :"
      Height          =   1455
      Left            =   3480
      TabIndex        =   6
      Top             =   960
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "    TRANSACCION"
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton CmdTransa 
         Caption         =   "Imprimir Trans."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tabla a Imprimir :"
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "Imprime B/D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   0
         Top             =   960
         Width           =   1455
      End
      Begin VB.ComboBox LIST_TABLAS 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Retornar"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Imprimir Tablas de  :   BDATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   2  'Blackness
      X1              =   0
      X2              =   5400
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "FrmImprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdTransa_Click()
On Error GoTo SALE
Dim RUTA, wconect, wname As String

crystal1.Connect = PUB_ODBC
crystal1.ReportFileName = PUB_RUTA_REPORTE & "transa.rpt"
crystal1.Action = 1
Exit Sub
SALE:
 MsgBox Err.Description, 48, Pub_Titulo



End Sub

Private Sub Command1_Click()
On Error GoTo sal
Dim MISIZE As Currency
Dim WT As String
Dim a, i As Integer

CONECT_TAB LIST_TABLAS.text
a = 0
t = 2
suma = 0
Gen_llave.Requery
pub_mensaje = "Nº de Campos : " & Gen_llave.rdoColumns.count & "  ¿Desea Imprimir los Campos de Tabla ... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbYes Then   ' El usuario eligió
Else
    Exit Sub
End If
LKCHEK = True
IMPRI_CAMPOS
Exit Sub
sal:
MsgBox Err.Number & Err.Description
End Sub

Private Sub Command2_Click()
Unload FrmImprime
End Sub



Public Sub CONECT_TAB(tabla As String)
On Error GoTo ERRO
pub_cadena = "SELECT * FROM "
pub_cadena = pub_cadena + tabla
Set LLAVE = CN.CreateQuery("", pub_cadena)
Set Gen_llave = LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
CORTA = "PASA"
Exit Sub
ERRO:
 'If CORTA <> "ERROR" Then
   MsgBox "NO EXISTE  TABLA ..!"
   CORTA = "ERROR"
 'End If
End Sub

Private Sub Form_Load()
Dim cad As String
For t = 0 To CN.rdoTables.count - 1
 LIST_TABLAS.AddItem CN.rdoTables(t).Name
Next
LIST_TABLAS.text = LIST_TABLAS.List(0)

End Sub



Public Function tipo(Codigo As Integer) As String
Dim VAR As String
Select Case Codigo
Case 1
 VAR = "CHAR"
Case 2
 VAR = "NUMERIC"
Case 3
 VAR = "DECIMAL"
Case 4
 VAR = "INTEGER"
Case 5
 VAR = "SMALLINT"
Case 6
 VAR = "FLOAT"
Case 7
 VAR = "REAL"
Case 8
 VAR = "DOUBLE"
Case 9
 VAR = "DATE"
Case 10
 VAR = "TIME"
Case 11
 VAR = "DATE-TIME"
Case 12
 VAR = "VARCHAR"
Case -1
 VAR = "LONGVARCHAR"
Case -2
 VAR = "BINARY"
Case -3
 VAR = "VARBINARY"
Case -4
 VAR = "LONGVARBINARY"
Case -5
 VAR = "BIGINT"
Case -6
 VAR = "TINYINT"
Case -7
 VAR = "rdTypeBIT"
End Select
tipo = VAR
End Function

Private Sub LIST_TABLAS_GotFocus()
'SendKeys "%{UP}"
End Sub



Public Sub IMPRI_CAMPOS()
Dim MISIZE As Currency
Dim WT As String
Dim a, i As Integer
Dim WNUM As String * 3
Dim WCAMPO As String * 20
Dim WTIPO As String * 10
Dim WSIZE As String * 10
Dim RUTA
RUTA = PUB_RUTA_REPORTE & Trim(FrmImprime.LIST_TABLAS.text) & ".txt"
'Draft 10cpi
Close #1
If LKCHEK Then Open RUTA For Output As #1
cuenta = cuenta + 1
WT = "TABLA : " & Mid(Gen_llave.Name, 15)
If LKCHEK Then Print #1, Spc(5); WT Else Printer.Print Spc(5); WT
If LKCHEK Then Print #1, " " Else Printer.Print ; " "
If LKCHEK Then Print #1, " " Else Printer.Print " "
If Not LKCHEK Then Printer.FontSize = 8

For i = 0 To Gen_llave.rdoColumns.count - 1
 a = a + 1
 WNUM = a - 1
 WCAMPO = Gen_llave.rdoColumns(i).Name
 WTIPO = tipo(Gen_llave.rdoColumns(i).Type)
 WSIZE = Gen_llave.rdoColumns(i).Size
' If Trim(WTIPO) = "NUMERIC" Then MsgBox WSIZE
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 5 Then WSIZE = "1,0"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 6 Then WSIZE = "2,0"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 7 Then WSIZE = "3,0"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 8 Then WSIZE = "4,0"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 9 Then WSIZE = "5,2"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 11 Then WSIZE = "7,2"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 12 Then WSIZE = "8,0"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 13 Then WSIZE = "9,0"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 15 Then WSIZE = "11,2"
 If Trim(WTIPO) = "NUMERIC" And Trim(WSIZE) = 17 Then WSIZE = "13,4"
 
   
'  If LKCHEK Then Print #1, Spc(5); WCAMPO Else: Printer.Print Spc(5); WCAMPO

  If LKCHEK Then Print #1, Spc(5); WNUM; ".-"; Spc(1); WCAMPO; Spc(1); WTIPO; Spc(1); WSIZE _
  Else Printer.Print Spc(5); WNUM; ".-"; Spc(1); WCAMPO; Spc(1); WTIPO; Spc(1); WSIZE
  
Next i
If LKCHEK Then Print #1, " " Else Printer.Print ; " "
If LKCHEK Then Print #1, Spc(5); "Total :."; a Else Printer.Print ; Spc(5); "Total : "; a
Close #1
WR_PAG = 999
If LKCHEK Then DoEvents: frmeditor.Show 1 Else Printer.EndDoc

Exit Sub

End Sub

Private Sub LIST_TABLAS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Command1.SetFocus
End If

End Sub
