Attribute VB_Name = "modSunat"
' Hecho   : por James
' Fecha   : 16/04/2015

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H100000
Private Const WAIT_INFINITE = -1&

Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function ShellAndWait(ByVal sPath As String, ByVal winStyle As VbAppWinStyle, Optional sTiempo As Long) As Boolean

    Dim procID As Long
    Dim procHandle As Long

    ' Start the program.
    On Error GoTo ShellError
    procID = Shell(sPath, vbHide)
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    procHandle = OpenProcess(SYNCHRONIZE, 0, procID)
    If procHandle <> 0 Then
        WaitForSingleObject procHandle, IIf(sTiempo = 0, WAIT_INFINITE, sTiempo)
        CloseHandle procHandle
    End If

    ' Reappear.
    ShellAndWait = True
    Exit Function

ShellError:
    ShellAndWait = False
End Function

Public Function GetShortDir(Nombre As String) As String
       Dim buffer As String
       buffer = String(255, 0)
       Call GetShortPathName(Nombre, buffer, 255)
       GetShortDir = Replace(buffer, Chr(0), vbNullString)
'    End If
End Function

Function GetDirTemp() As String
    If Environ$("temp") <> vbNullString Then
       GetDirTemp = Environ$("tmp")
    End If
End Function

Sub Descargar(URL As String)
  On Error GoTo Cualquiera
'    Me.MousePointer = vbHourglass

    Call URLDownloadToFile(0, URL, GetDirTemp & "\sunat.tmp", 0, 0)
'    Call URLDownloadToFile(0, URL, "c:\sunat.txt", 0, 0)

'    Me.MousePointer = vbDefault

    Exit Sub
Cualquiera:
'        Habilitar False
'        Limpiar
        MsgBox "No responde el servicio de la SUNAT", vbCritical, "Error"
End Sub

