VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#12.1#0"; "CODEJO~1.OCX"
Begin VB.Form frmProgramacionMain 
   BackColor       =   &H80000010&
   Caption         =   "Listado de Programaciones"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   270
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin XtremeCalendarControl.CalendarControl ccDatos 
      Height          =   5490
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   7575
      _Version        =   786433
      _ExtentX        =   13361
      _ExtentY        =   9684
      _StockProps     =   64
      ShowCaptionBar  =   -1  'True
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5910
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin XtremeCalendarControl.DatePicker dpMeses 
      Height          =   4335
      Left            =   9120
      TabIndex        =   0
      Top             =   420
      Width           =   2415
      _Version        =   786433
      _ExtentX        =   4260
      _ExtentY        =   7646
      _StockProps     =   64
      ShowNoneButton  =   0   'False
      ShowWeekNumbers =   -1  'True
      MaxSelectionCount=   1
      TextTodayButton =   "Hoy"
      VisualTheme     =   3
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      Caption         =   "CALENDARIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   80
      Width           =   1515
   End
   Begin VB.Label lblFecha 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000010&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   80
      Width           =   3555
   End
   Begin VB.Menu mnuMain 
      Caption         =   "MenuContextual"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuMEditarContrato 
         Caption         =   "Editar Contrato"
         Begin VB.Menu mnuAbrir 
            Caption         =   "Abrir"
         End
         Begin VB.Menu mnuEliminar 
            Caption         =   "Eliminar"
         End
         Begin VB.Menu mnuFinalizar 
            Caption         =   "Finalizar"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuVer 
            Caption         =   "Visualizar"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu mnuAnular 
            Caption         =   "Anular"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuMNuevoContrato 
         Caption         =   "Nuevo Contrato"
         Begin VB.Menu mnuNuevoContrato 
            Caption         =   "Nueva Programación"
         End
      End
      Begin VB.Menu mnuLineaTiempo 
         Caption         =   "Linea de Tiempo"
         Begin VB.Menu mnuTimeLine 
            Caption         =   "60 Minutos"
            HelpContextID   =   60
            Index           =   1
         End
         Begin VB.Menu mnuTimeLine 
            Caption         =   "30 Minutos"
            HelpContextID   =   30
            Index           =   2
         End
         Begin VB.Menu mnuTimeLine 
            Caption         =   "15 Minutos"
            HelpContextID   =   15
            Index           =   3
         End
         Begin VB.Menu mnuTimeLine 
            Caption         =   "10 Minutos"
            HelpContextID   =   10
            Index           =   4
         End
         Begin VB.Menu mnuTimeLine 
            Caption         =   "5 Minutos"
            HelpContextID   =   5
            Index           =   5
         End
      End
   End
End
Attribute VB_Name = "frmProgramacionMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private vCarga      As Boolean

Private vBusca      As Boolean

Public contextEvent As CalendarEvent

Private vF1         As Date

Private vF2         As Date

Private Sub ccDatos_ContextMenu(ByVal x As Single, ByVal y As Single)

    Dim HitTest As CalendarHitTestInfo

    Set HitTest = Me.ccDatos.ActiveView.HitTest
    
    If Not HitTest.ViewEvent Is Nothing Then
        Set contextEvent = HitTest.ViewEvent.Event
        Me.PopupMenu mnuMEditarContrato
        Set contextEvent = Nothing
        '    ElseIf (HitTest.HitCode = xtpCalendarHitTestDayViewTimeScale) Then
        '        Me.PopupMenu mnuContextTimeScale
        '        Me.PopupMenu mnuc
    ElseIf HitTest.HitCode = xtpCalendarHitTestDayViewTimeScale Then
        Me.PopupMenu mnuLineaTiempo
    Else
        
        Me.PopupMenu mnuMNuevoContrato
    End If

End Sub

Private Sub CargarContratos(vFechaIni As Date, vFechaFin As Date)
    'Dim evento As CalendarEvent
    'If vBUSCA Then
    Me.ccDatos.DataProvider.RemoveAllEvents
    
    'Me.ccDatos.DataProvider.DeleteEvent(

    If Me.ccDatos.ViewType = xtpCalendarMonthView Then
    End If

    Dim oRS As ADODB.Recordset
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SPPROGRAMACION_LIST"
  
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
  
    'oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@XFECHAS", adBoolean, adParamInput, , 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@INI", adDBTimeStamp, adParamInput, , vFechaIni)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FIN", adDBTimeStamp, adParamInput, , vFechaFin)
    
    Set oRS = oCmdEjec.Execute

    Dim evento As CalendarEvent
     
    Dim oREC   As CalendarRecurrencePattern
     
    Dim ORSd   As ADODB.Recordset
    
    Set ORSd = oRS.NextRecordset
    
    Dim vSUBJECT As String

    Do While Not oRS.EOF
        vSUBJECT = ""
        Set evento = Me.ccDatos.DataProvider.CreateEvent
        evento.StartTime = oRS!ini
        evento.EndTime = oRS!fin
        evento.Subject = "TURNO " & oRS!turno
        
        Dim ORSdd As ADODB.Recordset
       
        ORSd.Filter = "IDPROGRAMACION=" & oRS!IDPROGRAMACION
        
        Do While Not ORSd.EOF
            vSUBJECT = vSUBJECT & ORSd!Cantidad & " - " & ORSd!PRODUCTO & vbCrLf
            ORSd.MoveNext
        Loop


        evento.CreateRecurrence
        Set oREC = evento.RecurrencePattern

        oREC.Options.RecurrenceType = xtpCalendarRecurrenceDaily
        oREC.DurationMinutes = 60
        oREC.Options.DailyEveryWeekDayOnly = False
        oREC.Options.DailyIntervalDays = 1
        oREC.DurationMinutes = DateDiff("n", oRS!ini, oRS!fin)

        oREC.StartTime = TimeValue(oRS!ini)
        oREC.StartDate = DateValue(oRS!ini)


        oREC.EndMethod = xtpCalendarPatternEndDate
        oREC.EndDate = DateValue(oRS!fin)

        evento.UpdateRecurrence oREC
        
        ORSd.Filter = ""
        ORSd.MoveFirst
        evento.Body = vSUBJECT
        'evento.Subject = evento.Subject + Space(10) + vSUBJECT
        evento.ScheduleID = oRS!IDPROGRAMACION

        '        If ors!ESTADO = "ANULADO" Then
        '            evento.label = 1
        '        ElseIf ors!ESTADO = "VIGENTE" Then
        '            evento.label = 3
        '        ElseIf ors!ESTADO = "FINALIZADO" Then
        '            evento.label = 2
        '        Else
        '            evento.label = 4
        '        End If

        'label 5 = naranja
        'label 3 = verde
        'label 2 = azul
        'label 1 = rojo
        
        Me.ccDatos.DataProvider.AddEvent evento
        
        oRS.MoveNext
    Loop

    'End If
End Sub

Private Sub ccDatos_DblClick()

    Dim eCONTRATO As CalendarHitTestInfo

    Set eCONTRATO = ccDatos.ActiveView.HitTest

    If Not eCONTRATO.ViewEvent Is Nothing Then
        Set contextEvent = eCONTRATO.ViewEvent.Event
    End If

If Not eCONTRATO.ViewEvent Is Nothing Then

    MostrarContrato eCONTRATO.ViewEvent.Event.ScheduleID
    End If
    'mnuAbrir_Click
End Sub





Private Sub ccDatos_IsEditOperationDisabled(ByVal OpParams As XtremeCalendarControl.CalendarEditOperationParameters, _
                                            bDisableOperation As Boolean)

    If OpParams.EventViews(0).Event.RecurrenceState <> xtpCalendarRecurrenceNotRecurring Then
        bDisableOperation = True
    End If

End Sub



Private Sub ccDatos_ViewChanged()

    Dim NroDias As Long

    'vBUSCA = True

    NroDias = Me.ccDatos.ActiveView.DaysCount

    If NroDias = 1 Then
        vF1 = Me.ccDatos.ActiveView.Days(0).Date
        vF2 = Me.ccDatos.ActiveView.Days(0).Date
    Else
        vF1 = Me.ccDatos.ActiveView.Days(0).Date
        vF2 = Me.ccDatos.ActiveView.Days(NroDias - 1).Date
    End If

    CargarContratos vF1, vF2
       
    If (NroDias = 1) Then
        Me.lblFecha.Caption = Format(Me.ccDatos.ActiveView.Days(0).Date, "Long Date")
    ElseIf (NroDias > 1) Then
        Me.lblFecha.Caption = Format(Me.ccDatos.ActiveView.Days(0).Date, "Long Date") & " - " & Format(Me.ccDatos.ActiveView.Days(NroDias - 1).Date, "Long Date")
    End If
    
End Sub



Private Sub Form_Load()
    vCarga = False

    'Me.ccDatos.Options.TooltipAddNewText = "Agregar Contrato"
    Me.ccDatos.Options.EnableAddNewTooltip = False
    Me.ccDatos.Options.EnableInPlaceEditEventSubject_AfterEventResize = False
    Me.ccDatos.Options.EnableInPlaceEditEventSubject_ByMouseClick = False
    Me.ccDatos.Options.EnableInPlaceEditEventSubject_ByTab = False
    Me.ccDatos.Options.EnablePrevNextEventButtons = False
    Me.ccDatos.DayView.TimeScale = 20 'cada 20 minutos
    
    Me.ccDatos.Options.EnableInPlaceCreateEvent = False
    
    
    Dim orsT As ADODB.Recordset
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPPROGRAMACION_TURNO_FILL"
    Set orsT = oCmdEjec.Execute(, Array(LK_CODCIA))
    
    If Not orsT.EOF Then
        Me.ccDatos.DayView.TimeScaleMinTime = IIf(IsNull(orsT!Min), #8:00:00 AM#, orsT!Min)
        Me.ccDatos.DayView.TimeScaleMaxTime = IIf(IsNull(orsT!max), #12:00:00 PM#, orsT!max)
    Else
        Me.ccDatos.DayView.TimeScaleMinTime = #8:00:00 AM#
        Me.ccDatos.DayView.TimeScaleMaxTime = #12:00:00 PM#
    End If
    
    
    
    
    Me.ccDatos.Options.WorkWeekMask = xtpCalendarDayMo_Fr Or xtpCalendarDaySaturday
    
    Me.ccDatos.EnableReminders True
    Me.dpMeses.AttachToCalendar Me.ccDatos
    Me.ccDatos.ViewType = xtpCalendarFullWeekView

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    Dim nHeight As Long, LabelWidth As Long
    
    nHeight = Height - Me.sbMain.Height * 3 - Me.ccDatos.Top - 12 * Screen.TwipsPerPixelY + 300
    
    LabelWidth = Me.ScaleWidth - Me.lblTitulo.Width - 100
    
    If nHeight < 0 Then Height = 0
    If LabelWidth < 0 Then LabelWidth = 0

    Me.dpMeses.Left = Me.ScaleWidth - Me.dpMeses.Width - 3
    Me.dpMeses.Height = nHeight
        
    Me.ccDatos.Move 0, Me.ccDatos.Top, Me.dpMeses.Left, nHeight
        
    Me.lblFecha.Move Me.lblTitulo.Width, Me.lblFecha.Top, LabelWidth, Me.lblFecha.Height
    
End Sub

Private Sub MostrarContrato(xIdContrato As Integer)
'''    frmContratos.vIDContrato = xIdContrato ' contextEvent.ScheduleID ' Me.lvData.SelectedItem.Text
'''    '    frmContratos.dtpInicio.Value = xFECHAINICIO 'contextEvent.StartTime ' Me.lvData.SelectedItem.SubItems(2)
'''    '    frmContratos.dtpTermino.Value = xFECHAFIN ' contextEvent.EndTime ' Me.lvData.SelectedItem.SubItems(3)
'''
'''    frmContratos.vNUEVO = False
'''
'''    If contextEvent.label = 1 Then
'''        frmContratos.EstaAnulado = True
'''    Else
'''        frmContratos.EstaAnulado = False
'''    End If
'''
'''    frmContratos.Show vbModal
'''
    frmProgramacion.xIDPROGRAMACION = xIdContrato
    frmProgramacion.VNuevo = False
    frmProgramacion.Show vbModal
    CargarContratos vF1, vF2
End Sub

Private Sub mnuAbrir_Click()
    MostrarContrato contextEvent.ScheduleID
End Sub

Private Sub mnuAnular_Click()

    If contextEvent.label = 1 Then
        MsgBox "EL contrato seleccionado ya se encuentra anulado.", vbInformation, Pub_Titulo

        Exit Sub

    ElseIf contextEvent.label = 2 Then 'FINALIZADO
        MsgBox "No se puede Anular un contrato Finalizado.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    If MsgBox("¿Desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
    
        Dim vTIENEPAGOS As Boolean
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCONTRATOS_VERIFICAMORTIZACIONES"
        oCmdEjec.CommandType = adCmdStoredProc
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , contextEvent.ScheduleID) ' Me.lvData.SelectedItem.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@TIENEDATOS", adBoolean, adParamOutput, , vTIENEPAGOS)
        oCmdEjec.Execute
        
        vTIENEPAGOS = oCmdEjec.Parameters(2).Value
        
        If vTIENEPAGOS Then
            If MsgBox("El Contrato Seleccionado posee pagos registrados." & vbCrLf & "¿Desea ingresar Las Credenciales de Seguridad?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
            
                frmClaveCaja.Show vbModal

                If frmClaveCaja.vAceptar Then

                    Dim vS As String

                    If VerificaPass(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                        AnularContratos
                    Else
                        MsgBox vS, vbCritical, Pub_Titulo
                    End If
                End If
            
            End If

        Else
            AnularContratos
        End If
        
    End If

End Sub

Private Function VerificaPassFinaliza(vUSUARIO As String, _
                                      vClave As String, _
                                      ByRef vMSN As String) As Boolean

    Dim orsPass As ADODB.Recordset

    Dim vtpass  As String, vPasa As Boolean

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPDEVUELVECLAVEAMORTIZACIONES2"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
    oCmdEjec.Execute

    'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
    vtpass = oCmdEjec.Parameters("@MSN").Value
    vPasa = oCmdEjec.Parameters("@PASA").Value
    vMSN = vtpass

    VerificaPassFinaliza = vPasa
End Function

Private Function VerificaPass(vUSUARIO As String, _
                              vClave As String, _
                              ByRef vMSN As String) As Boolean

    Dim orsPass As ADODB.Recordset

    Dim vtpass  As String, vPasa As Boolean

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPDEVUELVECLAVEAMORTIZACIONES"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@USUARIO", adVarChar, adParamInput, 10, vUSUARIO)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CLAVE", adVarChar, adParamInput, 10, vClave)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@MSN", adVarChar, adParamOutput, 200, 1)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@PASA", adBoolean, adParamOutput, , 1)
    oCmdEjec.Execute

    'If Not orsPass.EOF Then vtpass = Trim(orsPass!Clave)
    vtpass = oCmdEjec.Parameters("@MSN").Value
    vPasa = oCmdEjec.Parameters("@PASA").Value
    vMSN = vtpass

    VerificaPass = vPasa
End Function

Private Sub AnularContratos()
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPFINALIZARANULARCONTRATO"
            
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , contextEvent.ScheduleID) ' Me.lvData.SelectedItem.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ANULANDO", adBoolean, adParamInput, , 1)
    oCmdEjec.Execute
                
    'Me.lvData.SelectedItem.SubItems(8) = "ANULADO"
        
    MsgBox "Contrato Anulado Correctamente.", vbInformation, Pub_Titulo
    CargarContratos vF1, vF2
End Sub

Private Sub mnuEliminar_Click()

    If MsgBox("¿Desea continuar con la operación.?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPPROGRAMACION_ELIMINAR_FULL"
        
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDPROGRAMACION", adBigInt, adParamInput, , contextEvent.ScheduleID) ' Me.lvData.SelectedItem.Text)
        oCmdEjec.Execute
        
        MsgBox "Datos Almacenados Correctamente.", vbInformation, Pub_Titulo
        CargarContratos vF1, vF2
    End If

End Sub

Private Sub mnuFinalizar_Click()

    If contextEvent.label = 1 Then
        MsgBox "No se puede Finalizar un Contrato Anulado.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    If contextEvent.label = 2 Then
        MsgBox "No se puede finalizar un Contrato ya Finalizado.", vbInformation, Pub_Titulo

        Exit Sub

    End If

    If MsgBox("¿Desea continuar con la operación?", vbQuestion + vbYesNo, Pub_Titulo) = vbYes Then
    
        Dim vTIENEPAGOS As Boolean
    
        LimpiaParametros oCmdEjec
        oCmdEjec.CommandText = "SPCONTRATOS_VALIDAPAGOSCOMPLETOS"
        oCmdEjec.CommandType = adCmdStoredProc
    
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , contextEvent.ScheduleID) ' Me.lvData.SelectedItem.Text)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
        oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@COMPLETO", adBoolean, adParamOutput, , vTIENEPAGOS)
        oCmdEjec.Execute
        
        vTIENEPAGOS = oCmdEjec.Parameters(2).Value
        
        If Not vTIENEPAGOS Then
            MsgBox "No se puede finalizar el Contrato debido a que no esta completamente pagado.", vbInformation, Pub_Titulo

            Exit Sub

        End If
        
        frmClaveCaja.Show vbModal

        If frmClaveCaja.vAceptar Then

            Dim vS As String

            If VerificaPassFinaliza(frmClaveCaja.vUSUARIO, frmClaveCaja.vClave, vS) Then
                FinalizaContrato
            Else
                MsgBox vS, vbCritical, Pub_Titulo
            End If
        End If
      
    End If

End Sub

Private Sub FinalizaContrato()
        
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPFINALIZARANULARCONTRATO"
            
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , contextEvent.ScheduleID) ' Me.lvData.SelectedItem.Text)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@ANULANDO", adBoolean, adParamInput, , 0)
    oCmdEjec.Execute
                
    'Me.lvData.SelectedItem.SubItems(8) = "FINALIZADO"
    MsgBox "Contrato Finalizado Correctamente.", vbInformation, Pub_Titulo
    CargarContratos vF1, vF2
End Sub



Private Sub mnuNuevoContrato_Click()
''    frmContratos.VNuevo = True
''    frmContratos.dtpInicio.Value = Me.ccDatos.DayView.Selection.Begin
''    frmContratos.dtpTermino.Value = Me.ccDatos.DayView.Selection.End
''    frmContratos.Show vbModal
''    CargarContratos vF1, vF2

frmProgramacionOpcion.gNUEVO = True
frmProgramacionOpcion.gFecha = Me.ccDatos.ActiveView.Selection.Begin
frmProgramacionOpcion.Show vbModal
CargarContratos vF1, vF2

End Sub

Private Sub mnuTimeLine_Click(Index As Integer)
 
 
 
 Me.ccDatos.DayView.TimeScale = mnuTimeLine(Index).HelpContextID
End Sub

Private Sub mnuVer_Click()
    'VisualizarContrato  Me.lvData.SelectedItem.Text
    VisualizarContrato contextEvent.ScheduleID
End Sub

Private Sub VisualizarContrato(xIdContrato As Integer)

    On Error GoTo Ver

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions

    Dim crParamDef  As CRAXDRT.ParameterFieldDefinition

    Dim objCrystal  As New CRAXDRT.Application

    Dim RutaReporte As String

    Dim oUSER       As String, oCLAVE As String, oLOCAL As String

    'RutaReporte = "C:\Admin\Nordi\Comanda1.rpt"
    
    'DATOS COMPLEMENTARIOS
    Dim orsC        As ADODB.Recordset

    LimpiaParametros oCmdEjec
    oCmdEjec.CommandText = "SPDATOSCOMPLEMENTARIOSCONTRATOS"
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    
    Set orsC = oCmdEjec.Execute
    
    'RutaReporte = "d:\VISTACONTRATO.rpt"
    RutaReporte = Trim(orsC!RutaReporte) + "VISTACONTRATO.rpt"
    oUSER = orsC!usuario
    oCLAVE = orsC!Clave
    oLOCAL = orsC!LOCAL

    If VReporte Is Nothing Then VReporte = New CRAXDRT.Report
    
    Set VReporte = objCrystal.OpenReport(RutaReporte, 1)
    Set crParamDefs = VReporte.ParameterFields
    
    LimpiaParametros oCmdEjec
    oCmdEjec.CommandType = adCmdStoredProc
    oCmdEjec.CommandText = "SPVISUALIZARCONTRATO"
    'oCmdEjec.CommandText = "SpPrintComanda"

    Dim rsd As ADODB.Recordset

    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , xIdContrato)
    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)

    Set rsd = oCmdEjec.Execute
    
    

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "TITULO"
                crParamDef.AddCurrentValue "CONTRATO " & rsd!tipo & " " & oLOCAL & " - Nº - " & contextEvent.ScheduleID  ' Me.lvData.SelectedItem.Text
        End Select

    Next

  

    '    Dim RSS As ADODB.Recordset
    '
    '    LimpiaParametros oCmdEjec
    '
    '    oCmdEjec.CommandText = "SPVISUALIZARCONTRATO2"
    '
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@IDCONTRATO", adBigInt, adParamInput, , xIdContrato)
    '    oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
    '
    '    Set RSS = oCmdEjec.Execute

    'SUB REPORTE
    Dim VReporteS As New CRAXDRT.Report
    Dim VReporteA As New CRAXDRT.Report

    VReporte.Database.SetDataSource rsd, , 1  'lleno el objeto reporte

    Set VReporteS = VReporte.OpenSubreport("DETALLE")
    Set VReporteA = VReporte.OpenSubreport("AMORTIZACIONES")
    
    VReporte.OpenSubreport("DETALLE").Database.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", oUSER, oCLAVE
    'VReporte.OpenSubreport("DETALLE").Database.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", oUSER, "accesodenegado"
    
    
    
    Set crParamDefs = VReporteS.ParameterFields

    For Each crParamDef In crParamDefs

        Select Case crParamDef.ParameterFieldName

            Case "Pm-ado.idcontrato"
                crParamDef.AddCurrentValue xIdContrato

            Case "Pm-ado.CODCIA"
                crParamDef.AddCurrentValue LK_CODCIA
        End Select

    Next
    
    'VReporte.OpenSubreport("AMORTIZACIONES").Database.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", oUSER, "accesodenegado"
    VReporte.OpenSubreport("AMORTIZACIONES").Database.LogOnServer "p2sodbc.dll", "DSN_DATOS", "bdatos", oUSER, oCLAVE
       
    'VReporteS.Database.SetDataSource RSS, , 1
 
    'VReporteS.ReadRecords
    frmContratosReporte.crContrato.ReportSource = VReporte

    'frmContratosReporte.crContrato.Refresh
    frmContratosReporte.crContrato.ViewReport

    frmContratosReporte.Show
    Set VReporte = Nothing
    Set VReporteS = Nothing

    Exit Sub

Ver:
    MostrarErrores Err
End Sub
