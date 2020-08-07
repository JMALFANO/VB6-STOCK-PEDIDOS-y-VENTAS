VERSION 5.00
Begin VB.Form frmRegistrarCliente 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "NUEVO CLIENTE:"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "DATOS:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6375
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1320
         TabIndex        =   17
         Top             =   4080
         Width           =   4935
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   3600
         Width           =   4935
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   3000
         Width           =   4935
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   2400
         Width           =   4935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GUARDAR CLIENTE"
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
         Left            =   120
         TabIndex        =   9
         Top             =   4680
         Width           =   6135
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   1800
         Width           =   4935
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Comentarios:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Horarios:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Contacto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mail:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefono:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuit:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2400
         Width           =   975
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "- JuanMAPCh.-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmRegistrarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Enabled = False Then
Call ActualizarCliente
Else
Call RegistrarCliente
End If
End Sub

Sub LimpiarDatos()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub
Private Sub Text1_Change()
Label2.Caption = Text1.Text
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmPrincipal.Visible = True
End Sub

Sub RegistrarCliente()

Cliente.nombre = Text1.Text
Cliente.telefono = Text2.Text
Cliente.direccion = Text3.Text
Cliente.mail = Text4.Text
Cliente.cuit = Text5.Text
Cliente.contacto = Text6.Text
Cliente.horarios = Text7.Text
Cliente.comentarios = Text8.Text
'esto por ahora juanam
Cliente.saldo = 0
'esto por ahora juanam
Dim NombreCliente As String


Dim i As Integer
Dim NombreC As String
Dim aux2 As String
Dim CantClientes As Integer

If FileExist(App.Path & "\Clientes\Clientes.TXT", vbNormal) Then 'ya se registro algun cliente?
CantClientes = Val(GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "CantClientes"))
For i = 0 To CantClientes
NombreC = (GetVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "NombreCliente" & i))
If Cliente.nombre = NombreC Then
MsgBox "ya se encuentra registrado el cliente: " & NombreC
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
Call WriteVar(App.Path & "\Clientes\Clientes.TXT", "CLIENTES", "NombreCliente" & aux2, Cliente.nombre)

Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Nombre", Cliente.nombre)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Telefono", Cliente.telefono)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Direccion", Cliente.direccion)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Mail", Cliente.mail)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Cuit", Cliente.cuit)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Contacto", Cliente.contacto)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Horarios", Cliente.horarios)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Comentarios", Cliente.comentarios)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "saldo", Cliente.saldo)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "NumeroCliente", aux2)

MsgBox "Cliente: " & Cliente.nombre & " registrado correctamente."
Call LimpiarDatos
Unload Me
Call CargarClientes
frmPrincipal.Visible = True
End Sub

Sub ActualizarCliente()
Cliente.nombre = Text1.Text
Cliente.telefono = Text2.Text
Cliente.direccion = Text3.Text
Cliente.mail = Text4.Text
Cliente.cuit = Text5.Text
Cliente.contacto = Text6.Text
Cliente.horarios = Text7.Text
Cliente.comentarios = Text8.Text

Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Telefono", Cliente.telefono)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Direccion", Cliente.direccion)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Mail", Cliente.mail)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Cuit", Cliente.cuit)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Contacto", Cliente.contacto)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Horarios", Cliente.horarios)
Call WriteVar(App.Path & "\Clientes\" & Cliente.nombre & ".TXT", "DATOS", "Comentarios", Cliente.comentarios)

MsgBox "Cliente: " & Cliente.nombre & " actualizado correctamente."
Unload Me
frmPrincipal.Visible = True
End Sub

