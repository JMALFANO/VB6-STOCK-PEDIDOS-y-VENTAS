VERSION 5.00
Begin VB.Form frmPedidosCliente 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "PEDIDAS"
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4095
      Begin VB.ListBox envases 
         Height          =   2985
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2535
      End
      Begin VB.ListBox cantidad 
         Enabled         =   0   'False
         Height          =   2985
         Left            =   2640
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad:"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Envase:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      Caption         =   "JuanMAPCh.-"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmPedidosCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmPrincipal.Visible = True
End Sub

