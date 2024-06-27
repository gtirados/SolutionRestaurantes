VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPLE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exporta Registro de Ventas al PLE"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   7350
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar Importe"
      Height          =   480
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Generar txt"
      Height          =   480
      Left            =   5280
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox ComMes 
      Height          =   315
      ItemData        =   "frmPLE.frx":0000
      Left            =   2760
      List            =   "frmPLE.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   180
      Width           =   2055
   End
   Begin VB.TextBox txtAnio 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   195
      Width           =   855
   End
   Begin MSComCtl2.UpDown udAnio 
      Height          =   285
      Left            =   1576
      TabIndex        =   2
      Top             =   195
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2000
      BuddyControl    =   "txtAnio"
      BuddyDispid     =   196612
      OrigLeft        =   2520
      OrigTop         =   2280
      OrigRight       =   2760
      OrigBottom      =   4215
      Max             =   2100
      Min             =   2000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPORTE A DECLARAR:"
      Height          =   195
      Left            =   480
      TabIndex        =   8
      Top             =   1080
      Width           =   2040
   End
   Begin VB.Label lblResumen 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label lblMES 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MES:"
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   435
   End
   Begin VB.Label lblLabel1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AÑO:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   450
   End
End
Attribute VB_Name = "frmPLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Exportar_ListView(xDatos As ADODB.Recordset, PathArchivo As String)
    On Error GoTo errsub
    Dim Linea As String, x As Integer, i As Integer
      
    'Abrimos un archivo para guardar los datos del ListView
    Open PathArchivo For Output As #1
      
'''    'Recorremos los encabezados para guardar el caption
'''    For i = 1 To ListView1.ColumnHeaders.count
'''        Linea = Linea & ListView1.ColumnHeaders(i).Text & vbTab
'''    Next
'''    'Imprimimos la línea
'''    Print #1, Linea
      
      i = 1
    'recorremos cada Item y Subitem
    Do While Not xDatos.EOF
        Linea = Left(xDatos.Fields("DATA").Value, 20) & "M" & Right("00000" & CStr(i), 5) & "|" & Mid(xDatos.Fields("DATA").Value, 21, Len(xDatos.Fields("DATA").Value))
        'Imprimimos la linea
    Print #1, Linea
    xDatos.MoveNext
    i = i + 1
    Loop
    
'''    For i = 1 To ListView.ListItems.count
'''        'texto del Item
'''        Linea = ListView.ListItems(i) & vbTab
'''        'texto de los SubItems
'''        For x = 1 To ListView1.ColumnHeaders.count - 1
'''            Linea = Linea & ListView.ListItems.Item(i).SubItems(x) & vbTab
'''        Next
'''    'Imprimimos la linea
'''    Print #1, Linea
'''    Next
      
    'Cerramos
    Close
  
Exit Sub
errsub:
MsgBox Err.Description, vbCritical
  
End Sub

Private Sub CmdExportar_Click()
  LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "USP_PLE_VENTAS"

    Dim oRSr As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("anio", adInteger, adParamInput, , Me.txtAnio.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("sd", adTinyInt, adParamInput, , Me.ComMes.ListIndex + 1)
    
Set oRSr = oCmdEjec.Execute

Dim ORSf As New ADODB.Recordset
LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "USP_EMPRESA_RUC"
 oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
 
 Set ORSf = oCmdEjec.Execute

Call Exportar_ListView(oRSr, App.Path & "\LE" & ORSf.Fields(0).Value & Me.txtAnio.Text & Right("00" & CStr(Me.ComMes.ListIndex + 1), 2) & "00140100001111.txt")

MsgBox "Archivo generado correctamente.", vbInformation, Pub_Titulo

End Sub

Private Sub cmdMostrar_Click()
 LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "USP_PLE_VENTAS"

    Dim oRSr As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("anio", adInteger, adParamInput, , Me.txtAnio.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("sd", adTinyInt, adParamInput, , Me.ComMes.ListIndex + 1)
    
Set oRSr = oCmdEjec.Execute

Dim orsRes As New ADODB.Recordset
Set orsRes = oRSr.NextRecordset
If orsRes.EOF Then
    MsgBox "No hay data.", vbCritical, Pub_Titulo
    Exit Sub
    
End If
Me.lblResumen.Caption = FormatCurrency(orsRes.Fields(0).Value, 2)
End Sub

Private Sub Form_Load()
CentrarFormulario MDIForm1, Me
Me.txtAnio.Text = Year(Date)
Me.ComMes.ListIndex = Month(Date) - 1
End Sub

