VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton OptionTAceite 
      Caption         =   "ACEITE"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton OptionTRosca 
      Caption         =   "ROSCA"
      Height          =   255
      Left            =   3360
      TabIndex        =   11
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ACEITE"
      Height          =   255
      Left            =   1440
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ROSCA"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "AGREGAR"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AGREGAR"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ListBox List4 
      Height          =   2595
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   4800
      TabIndex        =   3
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   3960
      TabIndex        =   2
      Top             =   4080
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label total 
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   5160
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400040&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   11
      Left            =   4560
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   10
      Left            =   4080
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   9
      Left            =   3600
      Top             =   2280
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   8
      Left            =   4800
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   7
      Left            =   4320
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   6
      Left            =   3840
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   5
      Left            =   3360
      Top             =   1800
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF80FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   4
      Left            =   4800
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   3
      Left            =   4320
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   2
      Left            =   3840
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   375
      Index           =   1
      Left            =   3360
      Top             =   1320
      Width           =   375
   End
   Begin VB.Shape Shape12 
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   3480
      Top             =   3000
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   11
      Left            =   4560
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   10
      Left            =   4080
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   9
      Left            =   3600
      Top             =   2280
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   8
      Left            =   4800
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   7
      Left            =   4320
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   6
      Left            =   3840
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   5
      Left            =   3360
      Top             =   1800
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   4
      Left            =   4800
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   3
      Left            =   4320
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   2
      Left            =   3840
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   3360
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   2640
      X2              =   2640
      Y1              =   3840
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VENTAS AL-MA PLAST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TapaColor As String
Dim Acumulador As Single

Private Sub Command1_Click()
Dim valor As Single
Dim cuenta As Single

List1.AddItem Text1.Text & " envases de " & List4.Text
If List4.Text = "Laboratorio 1000" Then valor = 3.16
If List4.Text = "Laboratorio 500" Then valor = 2.73
If List4.Text = "Laboratorio 250" Then valor = 2.41
If List4.Text = "Laboratorio 125" Then valor = 2.18
List2.AddItem valor

cuenta = Text1.Text * valor
List3.AddItem cuenta
Acumulador = Acumulador + cuenta

total.Caption = Acumulador

End Sub

Private Sub Command2_Click()
If OptionTAceite.Value = False And OptionTRosca.Value = False Then
MsgBox "Seleccione Pico de TAPA."
Exit Sub
End If

Dim valor As Single
Dim cuenta As Single
List1.AddItem Text2.Text & " tapas de color " & Label2.Caption
If OptionTRosca.Value = True Then valor = 0.36
If OptionTAceite.Value = True Then valor = 0.49

List2.AddItem valor
cuenta = Text2.Text * valor
List3.AddItem cuenta
Acumulador = Acumulador + cuenta

total.Caption = Acumulador

End Sub

Private Sub Form_Load()
List4.AddItem "Laboratorio 1000"
List4.AddItem "Laboratorio 500"
List4.AddItem "Laboratorio 250"
List4.AddItem "Laboratorio 125"
End Sub

Private Sub Image1_Click(Index As Integer)
Dim GuardoIndex As Byte
GuardoIndex = Index
Select Case Index

Case 1
If OptionTRosca.Value = True Then TapaColor = "Blanca"
If OptionTAceite.Value = True Then TapaColor = "Verde"
Case 2
If OptionTRosca.Value = True Then TapaColor = "Negra"
If OptionTAceite.Value = True Then TapaColor = "Amarilla"
Case 3
If OptionTRosca.Value = True Then TapaColor = "Transparente S/P"
If OptionTAceite.Value = True Then TapaColor = "Negra"
Case 4
If OptionTRosca.Value = True Then TapaColor = "Rosa"
If OptionTAceite.Value = True Then TapaColor = "Roja"
Case 5
TapaColor = "Celeste"
Case 6
TapaColor = "Verde Manzana"
Case 7
TapaColor = "Rojo"
Case 8
TapaColor = "Azul"
Case 9
TapaColor = "Amarillo"
Case 10
TapaColor = "Naranja"
Case 11
TapaColor = "Violeta"

End Select
Shape12.BackColor = Shape1(GuardoIndex).BackColor
Label2.Caption = TapaColor
End Sub

Private Sub OptionTRosca_Click()
Dim i As Byte

If OptionTRosca.Value = True Then
Shape1(1).BackColor = vbWhite
Shape1(2).BackColor = vbBlack
Shape1(3).BackColor = &HC0C0C0
Shape1(4).BackColor = &HFF80FF

For i = 5 To 11
Shape1(i).Visible = True
Next i

End If
End Sub

Private Sub OptionTAceite_Click()
Dim i As Byte
If OptionTAceite.Value = True Then
Shape1(1).BackColor = &H4000&
Shape1(2).BackColor = vbYellow
Shape1(3).BackColor = vbBlack
Shape1(4).BackColor = vbRed

For i = 5 To 11
Shape1(i).Visible = False
Next i

End If
End Sub

