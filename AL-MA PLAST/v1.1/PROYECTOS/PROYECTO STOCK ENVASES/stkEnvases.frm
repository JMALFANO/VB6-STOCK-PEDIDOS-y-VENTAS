VERSION 5.00
Begin VB.Form stkEnvases 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Designed by : Juan Martín ALFANO. (011) 15-5577-9985"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Nuevo Pedido:"
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   3855
      Begin VB.OptionButton Option4 
         Caption         =   "u"
         Height          =   255
         Left            =   3360
         TabIndex        =   28
         Top             =   840
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         Caption         =   "b"
         Height          =   255
         Left            =   2880
         TabIndex        =   27
         Top             =   840
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         OLEDropMode     =   2  'Automatic
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Top             =   840
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GUARDAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1800
         TabIndex        =   21
         Text            =   "Clientes"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Cliente:"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pedidas:"
      Height          =   3615
      Left            =   4080
      TabIndex        =   12
      Top             =   1920
      Width           =   3615
      Begin VB.CommandButton guardarabm 
         Caption         =   "GUARDAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   3375
      End
      Begin VB.ComboBox Comboabm 
         Height          =   315
         ItemData        =   "stkEnvases.frx":0000
         Left            =   2040
         List            =   "stkEnvases.frx":000D
         TabIndex        =   18
         Text            =   "Seleccione"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox abm 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   2055
      End
      Begin VB.ListBox List3 
         Enabled         =   0   'False
         Height          =   2205
         Left            =   2040
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Height          =   2205
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad:"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Envases:"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton Option2 
         Caption         =   "ACEITE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ROSCA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   2010
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label1 
         Caption         =   "Seleccionado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "PICO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label SELECCIONADO 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   2640
         Width           =   2175
      End
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   7680
      X2              =   4080
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label hay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "- STOCK ENVASES AL-MA PLAST-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   0
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "HAY  PEDIDAS  DISPONIBLES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   720
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   720
      Y2              =   1800
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   6120
      X2              =   6120
      Y1              =   720
      Y2              =   1800
   End
   Begin VB.Label pedidas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label disponibles 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "stkEnvases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cantidad As Integer
Dim CantString As String
Dim NombreCliente As String

Private Sub Combo1_GotFocus()
Text1.Text = ""
Call CargarClientesCombo
End Sub
Sub CargarClientesCombo()
Dim i As Integer
Dim CantClientes As String
CantClientes = Val(GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "CantClientes"))
Combo1.Clear ' limpio la lista pq sino me sigue mostrando lo que tenia cargado si ya cliqueò
For i = 0 To CantClientes
Combo1.AddItem GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "NombreCliente" & i)
Next i
End Sub
Sub ActualizoBolsas()
Envase.Stock = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Stock")
Envase.pedidas = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Pedidas")
SELECCIONADO.Caption = Envase.nombre
hay.Caption = Envase.Stock
pedidas.Caption = Envase.pedidas
disponibles.Caption = disponibles - Cantidad
Call BolsasDisponibles
End Sub
Sub GuardarCliente()
Dim i As Integer
Dim NombreC As String
Dim aux2 As String
Dim CantClientes As Integer

If FileExist(App.Path & "\Clientes\Clientes.TXT", vbNormal) Then 'ya se registro algun cliente?
CantClientes = Val(GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "CantClientes"))
For i = 0 To CantClientes
NombreC = (GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "NombreCliente" & i))
If NombreCliente = NombreC Then
'MsgBox "ya se encuentra registrado el cliente: " & NombreC
Exit Sub
End If
Next i
CantClientes = CantClientes + 1
aux2 = CantClientes
Else
CantClientes = 0
aux2 = CantClientes
End If

Call WriteVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "CantClientes", aux2)
Call WriteVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "NombreCliente" & aux2, Text1.Text)

'MsgBox " NOMBRE CLIENTE EN GUARDARCLIENTE: " & Text1.Text
End Sub
Sub GuardarPedido()
Dim i As Integer
Dim NombreC As String
Dim aux2 As String
Dim CantClientes As String
Dim EstaCliente As Boolean
Dim YaHabiaPedidas As Integer
Dim TeniaPedidas As String
Dim MP As String



CantClientes = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", "CLIENTES", "CantClientes")
If Not CantClientes = "" Then
YaHabiaPedidas = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Pedidas")
For i = 0 To CantClientes
NombreC = (GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", "CLIENTES", "" & i))

If NombreCliente = NombreC Then
EstaCliente = True
TeniaPedidas = GetVar(App.Path & "\Clientes\" & NombreC & ".TXT", "PEDIDOS", P & List1.ListIndex)
'MsgBox "ya se encuentra registrado el cliente EN EL ENVASE: " & NombreC
If Not TeniaPedidas <= 0 Then
MsgBox "Seleccione al cliente de la lista de pedidos."
Exit Sub
End If
End If
Next i

If Not EstaCliente = True Then
CantClientes = CantClientes + 1
aux2 = CantClientes
Else
aux2 = CantClientes '
'MsgBox "aca no te registro de nuevo al que acaba de hacer el pedido (:"
End If
Else
CantClientes = 0
aux2 = CantClientes
End If

If EstaCliente = False Then
Call WriteVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", "CLIENTES", "CantClientes", aux2)
Call WriteVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", "CLIENTES", "" & aux2, NombreCliente) 'jma
End If

MP = Val(GetVar(App.Path & "\Clientes\" & NombreCliente & ".TXT", "PEDIDOS", "M" & P))

If List1.ListIndex >= MP Then
Call WriteVar(App.Path & "\Clientes\" & NombreCliente & ".TXT", "PEDIDOS", "M" & P, List1.ListIndex)
'MsgBox "GUARDO ESTE VALOR PQ ES MAS ALTO" & List1.ListIndex
End If


Call WriteVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Pedidas", YaHabiaPedidas + CantString)
Call WriteVar(App.Path & "\Clientes\" & NombreCliente & ".TXT", "PEDIDOS", P & List1.ListIndex, CantString)
'Call LimpiarPedido 'aca j
Call LimpiarDefault
MsgBox "Pedido guardado."
Call ActualizarPedidos
End Sub
Private Sub Command1_Click()
If SELECCIONADO.Caption = "-" Then
MsgBox "Seleccione un envase."
Exit Sub
End If

If Text2.Text = "" Or Not IsNumeric(Text2.Text) Then
MsgBox "Asegurese de ingresar una cantidad."
Exit Sub
End If

If Text1.Text = "" And Combo1.Text <> "" And Not Combo1.Text = "Clientes" Then
NombreCliente = Combo1.Text
ElseIf Text1.Text <> "" And Combo1.Text = "Clientes" Then
NombreCliente = Text1.Text
Else
MsgBox "Asegurese de seleccionar un cliente."
Exit Sub
End If

If Option3.Value = True Then
Cantidad = Text2.Text
CantString = Cantidad
If CantString = 0 Then
MsgBox "La cantidad minima para realizar un pedido es 1 bolsón."
Exit Sub
End If
End If

If Option4.Value = True Then
Cantidad = Text2.Text / Envase.Cantidad
CantString = Cantidad
If CantString = 0 Then
MsgBox "La cantidad minima para realizar un pedido es 1 bolsón."
Exit Sub
End If
End If

'Call GuardarCliente
Call GuardarPedido

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub guardarabm_Click()
If SELECCIONADO.Caption = "-" Then
MsgBox "Seleccione un envase."
Exit Sub
End If

If abm.Text = "" Or Not IsNumeric(abm.Text) Then
MsgBox "Asegurese de ingresar una cantidad."
Exit Sub
End If


If Text1.Text = "STOCK" Then ' And Comboabm.Text = "Quitar" Then
Call GuardarstkEnvase(0, abm.Text, "stock")
Call ActualizarPedidos
Exit Sub
End If


If Comboabm.Text = "Seleccione" Then
MsgBox "Elija una opción."
Exit Sub
End If


If List2.Text = "" And Text1.Text <> "STOCK" Then
MsgBox "Debe seleccionar un cliente."
Exit Sub
End If

If List2.Text <> "" And Text1.Text <> "STOCK" Then

If Comboabm.Text = "Quitar" Then
Call GuardarstkEnvase(1, abm.Text, List2.Text)
End If
If Comboabm.Text = "Agregar" Then
Call GuardarstkEnvase(2, abm.Text, List2.Text)
End If
If Comboabm.Text = "Entregar" Then
Call GuardarstkEnvase(3, abm.Text, List2.Text)
End If
End If

Call ActualizarPedidos
End Sub

Private Sub BolsasDisponibles()
Dim BolsonesDisponibles As Integer

If Envase.pedidas = "" Then Envase.pedidas = 0
BolsonesDisponibles = Envase.Stock - Envase.pedidas

If BolsonesDisponibles > 0 Then disponibles.ForeColor = vbGreen
If BolsonesDisponibles < 0 Then disponibles.ForeColor = vbRed
If BolsonesDisponibles = 0 Then disponibles.ForeColor = vbBlack

disponibles.Caption = BolsonesDisponibles
End Sub
Private Sub LimpiarDefault()
Combo1.Text = "Clientes"
Text2.Text = ""
Option4.Value = False
Option3.Value = True
List2.Clear
List3.Clear
Label5.Caption = "Cliente:"
Label6.Caption = "Cantidad:"
abm.Text = ""
Text1.Text = ""
Comboabm.Text = "Seleccione"
hay.Caption = "0"
pedidas.Caption = "0"
disponibles.Caption = "0"
If disponibles.ForeColor = vbRed Or disponibles.ForeColor = vbGreen Then disponibles.ForeColor = vbBlack

End Sub
Private Sub List1_Click()
'Call LimpiarDefault ' pa que no muestre lo que ya cargo antes
Envase.nombre = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Nombre")
'Envase.Gramaje = GetVar(App.Path & "\stkEnvases\Rosca\R" & List1.ListIndex & ".TXT", "R" & List1.ListIndex, "Gramaje")
Envase.Pico = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Pico")
Envase.Cantidad = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Cantidad")
'Envase.Stock = GetVar(App.Path & "\stkEnvases\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Stock")
'Envase.pedidas = GetVar(App.Path & "\stkEnvases\" & Pico & "\" & P & List1.ListIndex & ".TXT", P & List1.ListIndex, "Pedidas")


Call ActualizarPedidos
End Sub
 Sub ActualizarPedidos()
 Call LimpiarDefault ' pa que no muestre lo que ya cargo antes
Dim i As Integer
Dim CantClientes As String
Dim ENVASESCLIENTE As String
Dim NombreCliente As String

 CantClientes = GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", "CLIENTES", "CantClientes")
 If CantClientes = "" Then
 Call ActualizoBolsas
 ' no hay clientes, pero mostrame que bolsa seleccione.
 Exit Sub
 End If
 For i = 0 To CantClientes
NombreCliente = (GetVar(App.Path & "\" & Pico & "\" & P & List1.ListIndex & ".TXT", "CLIENTES", "" & i))
ENVASESCLIENTE = (GetVar(App.Path & "\Clientes\" & NombreCliente & ".TXT", "PEDIDOS", P & List1.ListIndex))

'If NombreCliente = "" Then Exit For ' si no hay  clientes, con esto evito un error.
If ENVASESCLIENTE <> 0 Then 'aca no muestro si tiene cero bolsas pedidas
List2.AddItem NombreCliente
List3.AddItem ENVASESCLIENTE
End If
Next i

Call ActualizoBolsas

 End Sub
Private Sub List2_Click()
Label5.Caption = "Cliente: " & List2.Text
Label6.Caption = "Cantidad: " & List3.List(List2.ListIndex)
End Sub

Private Sub Option1_Click()
Call LimpiarDefault
Label5.Caption = "Cliente:"
SELECCIONADO.Caption = "-"
AceiteSeleccionado = False
RoscaSeleccionado = True
Pico = "Rosca"
P = "R"
Call CargarEnvases
End Sub
Private Sub Option2_Click()
Call LimpiarDefault
Label5.Caption = "Cliente:"
SELECCIONADO.Caption = "-"
RoscaSeleccionado = False
AceiteSeleccionado = True
Pico = "Aceite"
P = "A"
Call CargarEnvases
End Sub

Private Sub Text1_Click()
Combo1.Text = "Clientes"
End Sub
