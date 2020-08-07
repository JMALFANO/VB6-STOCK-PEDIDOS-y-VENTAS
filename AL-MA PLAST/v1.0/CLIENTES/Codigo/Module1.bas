Attribute VB_Name = "Module1"
Option Explicit


Public Type Clientes
 NOMBRE As String
 direccion As String
 telefono As String
 mail As String
 cuil As String
 saldo As String
End Type

Public Cliente As Clientes

Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal CrcKey As Long, ByVal CrcString As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Sub CargarClientes()
Dim CantClientes As Integer
Dim i As Integer


If FileExist(App.Path & "\Clientes\Clientes.TXT", vbNormal) Then 'ya se registro algun cliente?
CantClientes = Val(GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "CantClientes"))

frmPrincipal.List1.Clear

For i = 0 To CantClientes
Cliente.NOMBRE = GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "NombreCliente" & i)
frmPrincipal.List1.AddItem Cliente.NOMBRE
Next i

End If
End Sub
Sub WriteVar(file As String, Main As String, Var As String, Value As String) ' Funcion para crear archivos .txt .ini etc.
writeprivateprofilestring Main, Var, Value, file
End Sub
Function GetVar(file As String, Main As String, Var As String) As String ' Funcion para leer los archivos .txt .ini etc.
Dim sSpaces As String
sSpaces = Space$(5000)
getprivateprofilestring Main, Var, "", sSpaces, Len(sSpaces), file
GetVar = RTrim(sSpaces)
GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Function FileExist(file As String, FileType As VbFileAttribute) As Boolean ' Funcion para comprobar la existencia de archivo.

FileExist = Len(Dir$(file, FileType))

End Function
