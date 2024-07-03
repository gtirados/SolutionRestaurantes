Attribute VB_Name = "modSecurity"
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
 Public xParametros As New cParametros
Public oConec    As ADODB.Connection

Public Function getNombrePC() As String
  '-- Funcion auxiliar que devuelve el nombre del equipo llamando al API
  getNombrePC = Space$(260)
  GetComputerName getNombrePC, Len(getNombrePC)
  getNombrePC = Left$(getNombrePC, InStr(getNombrePC, vbNullChar) - 1)
End Function

Public Sub Conectar()

    'LK_FECHA_DIA = "27/10/2014"
    'LK_CODUSU = "ADMIN"
    Dim ss As String
    
    If oConec Is Nothing Then
        Set oConec = New ADODB.Connection
        oConec.CursorLocation = adUseClient
        oConec.CommandTimeout = 500
        'oConec.ConnectionTimeout = 500
        oConec.ConnectionString = xParametros.CadenaConexion
        oConec.Open
        
    Else

        If oConec.State = 0 Then
            oConec.CursorLocation = adUseClient
            oConec.CommandTimeout = 500
            oConec.ConnectionString = xParametros.CadenaConexion
            oConec.Open
            
        End If
    End If
   
End Sub


Public Function getUsuarioWindows() As String

    Dim lngRet      As Long

    Dim strUserName As String

    Dim strBuffer   As String * 25

    On Error GoTo ErrorHandler

    lngRet = GetUserName(strBuffer, 25)

    strUserName = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)

    getUsuarioWindows = strUserName

    Exit Function

ErrorHandler:

    MsgBox Err.Description

End Function
