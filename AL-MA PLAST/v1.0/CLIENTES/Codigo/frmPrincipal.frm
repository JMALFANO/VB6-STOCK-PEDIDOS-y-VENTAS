VERSION 5.00
Begin VB.Form frmPrincipal 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "REGISTRAR NUEVO CLIENTE"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      Caption         =   "CLIENTES:"
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BUSCAR"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   3765
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   4695
      End
   End
   Begin VB.CommandButton ventas 
      Caption         =   "VENTAS"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton pedidos 
      Caption         =   "PEDIDOS"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton datos 
      Caption         =   "DATOS"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label NOMBRE 
      Alignment       =   2  'Center
      Caption         =   "AL-MA PLAST"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command4_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
MsgBox "Desarrollando."
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Me.Visible = False
frmRegistrarCliente.Show

End Sub

Private Sub datos_Click()

If List1.Text = "" Then
MsgBox "Asegurese de seleccionar un cliente."
Exit Sub
End If

Me.Visible = False
frmDatosCliente.Show

Cliente.nombre = GetVar(App.Path & "\Clientes\" & frmPrincipal.List1.Text & ".TXT", "DATOS", "Nombre")
frmDatosCliente.nombre.Caption = Cliente.nombre

Cliente.telefono = GetVar(App.Path & "\Clientes\" & frmPrincipal.List1.Text & ".TXT", "DATOS", "Telefono")
frmDatosCliente.telefono.Caption = Cliente.telefono

Cliente.direccion = GetVar(App.Path & "\Clientes\" & frmPrincipal.List1.Text & ".TXT", "DATOS", "Direccion")
frmDatosCliente.direccion.Caption = Cliente.direccion

Cliente.mail = GetVar(App.Path & "\Clientes\" & frmPrincipal.List1.Text & ".TXT", "DATOS", "Mail")
frmDatosCliente.mail.Caption = Cliente.mail

Cliente.cuil = GetVar(App.Path & "\Clientes\" & frmPrincipal.List1.Text & ".TXT", "DATOS", "Cuil")
frmDatosCliente.cuil.Caption = Cliente.cuil

Cliente.saldo = Val(GetVar(App.Path & "\Clientes\" & frmPrincipal.List1.Text & ".TXT", "DATOS", "Saldo"))

If Cliente.saldo > 0 Then frmDatosCliente.saldo.ForeColor = vbRed
If Cliente.saldo < 0 Then frmDatosCliente.saldo.ForeColor = vbGreen

frmDatosCliente.saldo.Caption = Cliente.saldo
End Sub

Private Sub Form_Load()
Call CargarClientes

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub List1_Click()
Cliente.nombre = GetVar(App.Path & "\Clientes\" & List1.Text & ".TXT", "DATOS", "Nombre")
nombre.Caption = Cliente.nombre
End Sub

Private Sub pedidos_Click()
Cliente.nombre = GetVar(App.Path & "\Clientes\" & frmPrincipal.List1.Text & ".TXT", "DATOS", "Nombre")
frmPedidosCliente.nombre.Caption = Cliente.nombre

Dim EnvasesRosca As String
Dim EnvasesAceite As String
Dim NombreEnvase As String
Dim CantidadEnvase As String
Dim i As Integer
Dim y As Integer

Me.Visible = False
frmPedidosCliente.Show

EnvasesRosca = Val(GetVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "PEDIDOS", "MR"))
If EnvasesRosca >= O Then
For i = 0 To EnvasesRosca

CantidadEnvase = Val(GetVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "PEDIDOS", "R" & i))

If Not CantidadEnvase = 0 Then
NombreEnvase = (GetVar(App.Path & "\Rosca\Rosca.TXT", "R" & i, "R" & i))
frmPedidosCliente.envases.AddItem NombreEnvase
frmPedidosCliente.cantidad.AddItem CantidadEnvase
End If
Next i
End If

EnvasesAceite = Val(GetVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "PEDIDOS", "MA"))
If EnvasesAceite >= O Then
For y = 0 To EnvasesAceite

CantidadEnvase = Val(GetVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "PEDIDOS", "A" & y))

If Not CantidadEnvase = 0 Then
NombreEnvase = (GetVar(App.Path & "\Aceite\Aceite.TXT", "A" & y, "A" & y))
frmPedidosCliente.envases.AddItem NombreEnvase
frmPedidosCliente.cantidad.AddItem CantidadEnvase
End If
Next y
End If
End Sub

Private Sub ventas_Click()
MsgBox "Desarrollando."
End Sub
