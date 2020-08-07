Attribute VB_Name = "Module1"
Option Explicit
Public Flag As Boolean

Public Type Clientes
 NOMBRE As String
 direccion As String
 telefono As String
 mail As String
 cuit As String
 saldo As String
 contacto As String
 horarios As String
 comentarios As String
End Type

Public Cliente As Clientes

Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal CrcKey As Long, ByVal CrcString As String) As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Sub CargarClientes()
Dim CantClientes As Integer
Dim i As Integer
Dim y As Integer
Dim item As ListItem
Dim EnvasesRosca As String
Dim Stock As String 'estadopedido
Dim PuedoEntregar As Integer 'puedoentregar
Dim CantidadPedidos As String 'CANTIDADENVASE
Dim HayPedidas As String 'pedidastotal
'Dim NoPuedo As Boolean

If FileExist(App.Path & "\Clientes\Clientes.TXT", vbNormal) Then 'ya se registro algun cliente?
CantClientes = Val(GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "CantClientes")) 'CUANTOS?

frmPrincipal.ListView1.ListItems.Clear ' BORRAMOS LO Q YA ESTABA EN LA LISTA

frmPrincipal.ListView1.View = lvwReport
For i = 0 To CantClientes
Cliente.NOMBRE = GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "NombreCliente" & i) 'DAME EL NOMBRE DEL CLIENTE


    With frmPrincipal.ListView1.ListItems
        ' agrega algunos items
        Set item = .Add(, , Cliente.NOMBRE) 'CARGAMELO EN LA LISTA
              
        ' vista en modo reporte
       
    End With
    
EnvasesRosca = GetVar(App.Path & "\Clientes\" & Cliente.NOMBRE & ".TXT", "PEDIDOS", "MR") ' QUE ENVASES TIENE PEDIDOS?

If Not EnvasesRosca = "" Then 'TIENE ENVASES PEDIDOS?
For y = 0 To EnvasesRosca

CantidadPedidos = Val(GetVar(App.Path & "\Clientes\" & Cliente.NOMBRE & ".TXT", "PEDIDOS", "R" & y))
Stock = Val(GetVar(App.Path & "\Rosca\R" & y & ".TXT", "R" & y, "Stock"))
HayPedidas = Val(GetVar(App.Path & "\Rosca\R" & y & ".TXT", "R" & y, "Pedidas"))
PuedoEntregar = Stock - HayPedidas ') - CantidadPedidos)

If Not CantidadPedidos = 0 Then ' NO TIENE ENVASES PEDIDOS, MOSTRAME NEGRO..

If PuedoEntregar >= 0 Then
frmPrincipal.ListView1.ListItems(i + 1).ForeColor = vbGreen
Else
frmPrincipal.ListView1.ListItems(i + 1).ForeColor = vbRed
Exit For
End If

End If
PuedoEntregar = 0
      Next y
      
      End If
       
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

Function MostrarPedidos()

Dim EnvasesRosca As String
Dim EnvasesAceite As String
Dim NombreEnvase As String
Dim CantidadEnvase As String
Dim i As Integer
Dim y As Integer
Dim item As ListItem
Dim EstadoPedido As String
Dim Stock As String
Dim PuedoEntregar As Integer
Dim PedidasTotal As String

Cliente.NOMBRE = GetVar(App.Path & "\Clientes\" & frmPrincipal.ListView1.SelectedItem.Text & ".TXT", "DATOS", "Nombre")
frmPedidosCliente.NOMBRE.Caption = Cliente.NOMBRE

frmPedidosCliente.Show

frmPedidosCliente.ListView1.ListItems.Clear


frmPedidosCliente.ListView1.View = lvwReport

EnvasesRosca = Val(GetVar(App.Path & "\Clientes\" & Cliente.NOMBRE & ".TXT", "PEDIDOS", "MR"))
If EnvasesRosca >= 0 Then
For i = 0 To EnvasesRosca

CantidadEnvase = Val(GetVar(App.Path & "\Clientes\" & Cliente.NOMBRE & ".TXT", "PEDIDOS", "R" & i))
EstadoPedido = Val(GetVar(App.Path & "\Rosca\R" & i & ".TXT", "R" & i, "Stock"))
PedidasTotal = Val(GetVar(App.Path & "\Rosca\R" & i & ".TXT", "R" & i, "Pedidas"))
PuedoEntregar = EstadoPedido - (PedidasTotal - CantidadEnvase)
If Not CantidadEnvase = 0 Then
NombreEnvase = (GetVar(App.Path & "\Rosca\Rosca.TXT", "R" & i, "R" & i))

 With frmPedidosCliente.ListView1.ListItems
        ' agrega algunos items
        Set item = .Add(, , NombreEnvase)
        item.SubItems(1) = CantidadEnvase
        ' vista en modo reporte
       
    End With
If PuedoEntregar >= 0 Then frmPedidosCliente.ListView1.ListItems(i).ForeColor = vbGreen

If PuedoEntregar < 0 Then frmPedidosCliente.ListView1.ListItems(i + 1).ForeColor = vbRed

End If






Next i
End If

End Function

Sub Autocompletar(ListView As ListView, TBox As TextBox)
      
    ' variable para usar con el método FindItem que _
      permite buscar en el LV
    Dim item As ListItem
    Dim seleccion As Integer
      
    ' busca en el item, la cadena escrita en el textbox, si coincide _
      devuelve una referencia al item
    Set item = ListView.FindItem(TBox.Text, 0, , 1)
          
        ' verifica que el item no sea un valor nothing
        If Not item Is Nothing Then
            ' Muestra la selección pormas que no tenga el foco
            ListView.HideSelection = False
              
            ' desplaza la lista
            item.EnsureVisible
              
            ' selecciona el item
            item.Selected = True
  
            If Not Flag Then
                ' Almacena la posición de la selección en el textbox
                seleccion = TBox.SelStart
                  
                ' Asigna el texto completo del item encontrado
                TBox.Text = CStr(item)
                    If Not TBox.Text = vbNullString Then
                        ' posición de la selección
                        TBox.SelStart = seleccion
                        ' selecciona el texto
                        TBox.SelLength = Len(TBox.Text) - seleccion
                    End If
            End If
        Else
            ' Oculta la selección ya que no hay coincidencia
            ListView.HideSelection = True
        End If
  
End Sub



