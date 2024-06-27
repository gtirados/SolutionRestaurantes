VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private gServer As String
Private gDataBase As String
Private gUser As String
Private gPassword As String
Private gCadenaConexion As String
Private gCiaRegistro As String


Public Property Get CiaRegistro() As String
CiaRegistro = gCiaRegistro
End Property
Public Property Let CiaRegistro(xCC As String)
gCiaRegistro = xCC
End Property

Public Property Get CadenaConexion() As String
CadenaConexion = gCadenaConexion
End Property
Public Property Let CadenaConexion(xCC As String)
gCadenaConexion = xCC
End Property


Public Property Get Server() As String
 Server = gServer
End Property


Public Property Let Server(ByVal xserver As String)
 gServer = xserver
End Property


Public Property Get DataBase() As String
DataBase = gDataBase
End Property

Public Property Let DataBase(xDataBase As String)
gDataBase = xDataBase
End Property


Public Property Get User() As String
User = gUser
End Property

Public Property Let User(xUser As String)
gUser = xUser
End Property

Public Property Get Password() As String
Password = gPassword
End Property

Public Property Let Password(xPassword As String)
gPassword = xPassword
End Property

