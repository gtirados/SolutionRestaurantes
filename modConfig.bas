Attribute VB_Name = "modConfig"

'Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Sub EscribirINI(S As String, C As String, v As String, ArchINI As String)
   'S= Seccion a escribir
   'C=Clave a escribir
   'V=Valor a escribir
   'Archini=Archivo a leer
   'Escribo finalmente en el archivo INI
   WritePrivateProfileString S, C, v, ArchINI
End Sub

Public Function LeerIni(S As String, C As String, ArchivoINI As String) As String

   'S=Seccion de la cual leer
   'C=Clave a leer
   'Archivoini=Archivo a leer
   Dim VAR As String
   VAR = Space(128)
   Dim r As Long
   r = GetPrivateProfileString(S, C, "ERROR AL LEER INI", VAR, 128, ArchivoINI)
   LeerIni = Left(VAR, Len(VAR))
     LeerIni = Left(LeerIni, InStr(LeerIni, Chr(0)) - 1)
End Function

