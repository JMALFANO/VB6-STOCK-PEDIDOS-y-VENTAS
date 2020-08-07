VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
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
   Begin MSComctlLib.ListView ListView1 
      Height          =   3855
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6800
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "LISTADO CLIENTES"
         Object.Width           =   8116
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REGISTRAR NUEVO CLIENTE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
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
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton ventas 
      Caption         =   "VENTAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton pedidos 
      Caption         =   "PEDIDOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton datos 
      Caption         =   "DATOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   6
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

Private Sub Command2_Click()
Me.Visible = False
frmRegistrarCliente.Show
End Sub

Private Sub Text1_Change()
    ' autocompleta el Listview
    Call Autocompletar(ListView1, Text1)
End Sub
  
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
  
    Flag = False
    ' si se presionan el tecla de retroceso y la de borrar _
      no se autocompleta
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        Flag = True
    End If
  
End Sub
Private Sub datos_Click()

If ListView1.SelectedItem.Text = "" Then
MsgBox "Asegurese de seleccionar un cliente."
Exit Sub
End If

Me.Visible = False
frmDatosCliente.Show

Cliente.nombre = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Nombre")
frmDatosCliente.nombre.Caption = Cliente.nombre

Cliente.telefono = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Telefono")
frmDatosCliente.telefono.Caption = Cliente.telefono

Cliente.direccion = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Direccion")
frmDatosCliente.direccion.Caption = Cliente.direccion

Cliente.mail = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Mail")
frmDatosCliente.mail.Caption = Cliente.mail

Cliente.cuit = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Cuit")
frmDatosCliente.cuit.Caption = Cliente.cuit

Cliente.contacto = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Contacto")
frmDatosCliente.contacto.Caption = Cliente.contacto

Cliente.horarios = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Horarios")
frmDatosCliente.horarios.Caption = Cliente.horarios

Cliente.comentarios = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Comentarios")
frmDatosCliente.comentarios.Caption = Cliente.comentarios


Cliente.saldo = Val(GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Saldo"))

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
Private Sub ListVIEW1_Click()


Cliente.nombre = GetVar(App.Path & "\Clientes\" & ListView1.SelectedItem.Text & ".TXT", "DATOS", "Nombre")
nombre.Caption = Cliente.nombre
'Dim VARIABLE
'VARIABLE = ListView1.SelectedItem.Text
'MsgBox "He clicado el elemento " & VARIABLE
End Sub

Private Sub pedidos_Click()
Me.Visible = False
Call MostrarPedidos
End Sub

Private Sub ventas_Click()
MsgBox "Desarrollando."
End Sub






  
