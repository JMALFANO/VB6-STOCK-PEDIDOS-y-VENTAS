Attribute VB_Name = "Module1"
Option Explicit
Public RoscaSeleccionado As Boolean 'Utilizo para cargar la lista.
Public AceiteSeleccionado As Boolean 'Utilizo para cargar la lista.
Public Pico As String 'Rosca - Aceite
Public P As String 'R - A

Public Type Envases
 nombre As String
 Gramaje As String
 Pico As String
 Cantidad As String
 Stock As String
 pedidas As String
End Type

Public Type Clientes
 nombre As String
 direccion As String
 telefono As String
 cuil As String
 saldo As String
End Type


Public CantidadTotalEnvases As Byte 'Rosca o Aceite (for carga stock)
Public Envase As Envases
Public CLIENTE As Clientes
Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal CrcKey As Long, ByVal CrcString As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Sub CargarEnvases()
 
Dim i As Byte
Dim NombreEnvase As String
Dim NombreCliente As String
Dim CantidadEnvases As String

If RoscaSeleccionado Then
CantidadTotalEnvases = GetVar(App.Path & "\Rosca\Rosca.TXT", "ROSCA", "Cantidad")
End If

If AceiteSeleccionado Then
CantidadTotalEnvases = GetVar(App.Path & "\Aceite\Aceite.TXT", "ACEITE", "Cantidad")
End If


If stkEnvases.Visible = True Then
stkEnvases.List1.Clear 'LIMPIAMOS TODA LA LISTA SI CAMBIA DE ROSCA A ACEITE O INVERSA.
End If

For i = 0 To CantidadTotalEnvases

If RoscaSeleccionado Then
NombreEnvase = GetVar(App.Path & "\Rosca\R" & i & ".TXT", "R" & i, "Nombre")
End If

If AceiteSeleccionado Then
NombreEnvase = GetVar(App.Path & "\Aceite\A" & i & ".TXT", "A" & i, "Nombre")
End If



If stkEnvases.Visible = True Then
stkEnvases.List1.AddItem NombreEnvase
End If

Next i
End Sub
Public Sub GuardarstkEnvase(opcion As Byte, Cantidad As Integer, CLIENTE As String) '1 quitar '2 agregar ' 3 cliente
Dim Calculo As String
Dim NombreCliente As String
Dim CantidadClientes As String
Dim PedidasCliente As String
Dim PedidasTotal As Integer
Dim Stock As Integer
Dim i As Integer

'If CLIENTE <> "stock" Then

If opcion = 0 Then ' stock
Call WriteVar(App.Path & "\" & Pico & "\" & P & stkEnvases.List1.ListIndex & ".TXT", P & stkEnvases.List1.ListIndex, "Stock", Str(Cantidad))
MsgBox "Stock actualizado."
Exit Sub
End If

CantidadClientes = GetVar(App.Path & "\" & Pico & "\" & P & stkEnvases.List1.ListIndex & ".TXT", "CLIENTES", "CantClientes")

For i = 0 To CantidadClientes
NombreCliente = GetVar(App.Path & "\" & Pico & "\" & P & stkEnvases.List1.ListIndex & ".TXT", "CLIENTES", "" & i)
If NombreCliente = CLIENTE Then
PedidasCliente = GetVar(App.Path & "\Clientes\" & NombreCliente & ".TXT", "PEDIDOS", P & stkEnvases.List1.ListIndex)
Exit For
End If
Next i

PedidasTotal = GetVar(App.Path & "\" & Pico & "\" & P & stkEnvases.List1.ListIndex & ".TXT", P & stkEnvases.List1.ListIndex, "Pedidas")

If opcion = 1 Then ' quitar
Calculo = PedidasCliente - Cantidad
PedidasTotal = PedidasTotal - Cantidad
If PedidasTotal < 0 Then
MsgBox "La cantidad a quitar no puede ser mayor a la cantidad de pedidas."
Exit Sub
End If

End If


If opcion = 2 Then ' agregar
Calculo = PedidasCliente + Cantidad
PedidasTotal = PedidasTotal + Cantidad
End If



If opcion = 3 Then ' entregar
Stock = GetVar(App.Path & "\" & Pico & "\" & P & stkEnvases.List1.ListIndex & ".TXT", P & stkEnvases.List1.ListIndex, "Stock")
Stock = Stock - Cantidad
If Stock < 0 Then
MsgBox "La cantidad a entregar no puede ser mayor al stock disponible."
Exit Sub
End If

Calculo = PedidasCliente - Cantidad

If Calculo < 0 Then
MsgBox "La cantidad a entregar no puede ser mayor a la cantidad de pedidas."
Exit Sub
End If
PedidasTotal = PedidasTotal - Cantidad

Call WriteVar(App.Path & "\" & Pico & "\" & P & stkEnvases.List1.ListIndex & ".TXT", P & stkEnvases.List1.ListIndex, "Stock", Str(Stock))
End If

Call WriteVar(App.Path & "\" & Pico & "\" & P & stkEnvases.List1.ListIndex & ".TXT", P & stkEnvases.List1.ListIndex, "Pedidas", Str(PedidasTotal))
Call WriteVar(App.Path & "\Clientes\" & NombreCliente & ".TXT", "PEDIDOS", P & stkEnvases.List1.ListIndex, Calculo)

MsgBox "Actualizado."

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

