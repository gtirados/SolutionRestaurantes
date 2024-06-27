VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMesasMozo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Mozos a Mesas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7365
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   600
      Left            =   5880
      Picture         =   "frmMesasMozo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvMozos 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lvMesas 
      Height          =   3495
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOZOS ASIGNADOS:"
      Height          =   195
      Left            =   3600
      TabIndex        =   4
      Top             =   240
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MESAS:"
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmMesasMozo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGrabar_Click()

    If Me.lvMozos.ListItems.count = 0 Then
        MsgBox "No hay mozos para asignar", vbCritical, Pub_Titulo

        Exit Sub

    End If

    Dim xListaMozos As String

    xListaMozos = ""

    Dim f As Integer

    For f = 1 To Me.lvMesas.ListItems.count

        If Me.lvMesas.ListItems(f).Checked Then
            If Len(Trim(xListaMozos)) = 0 Then
                xListaMozos = Me.lvMesas.ListItems(f).Tag + ","
            Else
                xListaMozos = xListaMozos + Me.lvMesas.ListItems(f).Tag + ","
            End If
        
        End If

    Next

    On Error GoTo xGraba
    
    'VERIFICANDO SI ALGUNA MESA YA FUE ASIGNADA A OTRO MOSO
    
    Dim ORSv As ADODB.Recordset
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPVERIFICA_MESASbyMOZOS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MOZO", adDouble, adParamInput, , Me.lvMozos.SelectedItem.Tag)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MESAS", adVarChar, adParamInput, 4000, xListaMozos)
    
    Set ORSv = oCmdEjec.Execute
    
    If CBool(ORSv!Dato) Then
        If MsgBox("Hay mesas asignadas a otros usuarios." + vbCrLf + "¿Desea continuar?", vbQuestion + vbYesNo, Pub_Titulo) = vbNo Then Exit Sub
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPMESAMOZO_REGISTRAR"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MOZO", adDouble, adParamInput, , Me.lvMozos.SelectedItem.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MESAS", adVarChar, adParamInput, 4000, xListaMozos)
    
        oCmdEjec.Execute
        MsgBox "Datos Grabados correctamente.", vbInformation, Pub_Titulo
    Else
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPMESAMOZO_REGISTRAR"
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@codcia", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MOZO", adDouble, adParamInput, , Me.lvMozos.SelectedItem.Tag)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MESAS", adVarChar, adParamInput, 4000, xListaMozos)
    
        oCmdEjec.Execute
        MsgBox "Datos Grabados correctamente.", vbInformation, Pub_Titulo
    End If
    
    Exit Sub

xGraba:
    MsgBox Err.Description, vbCritical, Pub_Titulo

End Sub

Private Sub Form_Load()
ConfigurarLVS
    Dim orsM  As ADODB.Recordset

    Dim itemM As ListItem

   
CentrarFormulario MDIForm1, Me


   Dim oRsMZ  As ADODB.Recordset

    
    Me.lvMozos.ListItems.Clear

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SpListarMozos"
    Set oRsMZ = oCmdEjec.Execute(, LK_CODCIA)

    Do While Not oRsMZ.EOF

        With Me.lvMozos.ListItems.Add(, , Trim(oRsMZ!mozo))
            .Tag = Trim(oRsMZ!Codigo)
        End With

        '    Set ItemM = Me.lvMozo.ListItems.Add(, , Trim(oRsMozos!mozo))
        '    ItemM.Tag = Trim(oRsMozos!Codigo)
        oRsMZ.MoveNext
    Loop
    
    
     LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPMESASxMOZOS"
    Set orsM = oCmdEjec.Execute(, Array(LK_CODCIA, -1))

    Do While Not orsM.EOF

        With Me.lvMesas.ListItems.Add(, , Trim(orsM!mesa))
            .Tag = Trim(orsM!IDE)
        End With

        '    Set ItemM = Me.lvMozo.ListItems.Add(, , Trim(oRsMozos!mozo))
        '    ItemM.Tag = Trim(oRsMozos!Codigo)
        orsM.MoveNext
    Loop

End Sub

Private Sub ConfigurarLVS()

    With Me.lvMesas
        .Gridlines = True
        .LabelEdit = lvwManual
        .View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Mesa", 2500
          .CheckBoxes = True
    End With

   With Me.lvMozos
        .Gridlines = True
        .LabelEdit = lvwManual
      
        .View = lvwReport
        .FullRowSelect = True
        .ColumnHeaders.Add , , "Mozo", 3000
    End With
    
End Sub

Private Sub lvMozos_ItemClick(ByVal Item As MSComctlLib.ListItem)

   Dim oRsMZ  As ADODB.Recordset

    Dim itemM As ListItem
    Me.lvMesas.ListItems.Clear

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPMESASxMOZOS"
    Set oRsMZ = oCmdEjec.Execute(, Array(LK_CODCIA, Me.lvMozos.SelectedItem.Tag))

    Do While Not oRsMZ.EOF

        With Me.lvMesas.ListItems.Add(, , Trim(oRsMZ!mesa))
            .Tag = Trim(oRsMZ!IDE)
            .Checked = oRsMZ!ASIGNADO
        End With

        '    Set ItemM = Me.lvMozo.ListItems.Add(, , Trim(oRsMozos!mozo))
        '    ItemM.Tag = Trim(oRsMozos!Codigo)
        oRsMZ.MoveNext
    Loop
    
End Sub
