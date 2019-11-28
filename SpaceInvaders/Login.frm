VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   0  'None
   Caption         =   "ENTRAR EN CUENTA"
   ClientHeight    =   8820
   ClientLeft      =   6375
   ClientTop       =   615
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   Picture         =   "Login.frx":0000
   ScaleHeight     =   8820
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt4 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "Contraseña"
      Top             =   7680
      Width           =   6015
   End
   Begin VB.TextBox txt3 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "Nombre"
      Top             =   7080
      Width           =   6015
   End
   Begin VB.TextBox txt2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   8280
      Top             =   8400
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   585
      Left            =   1680
      Picture         =   "Login.frx":33955
      Top             =   8160
      Width           =   5130
   End
   Begin VB.Image Image5 
      Height          =   915
      Left            =   5040
      Picture         =   "Login.frx":36928
      Top             =   5760
      Width           =   2115
   End
   Begin VB.Image Image4 
      Height          =   930
      Left            =   1200
      Picture         =   "Login.frx":386BF
      Top             =   5760
      Width           =   3210
   End
   Begin VB.Image Image3 
      Height          =   1290
      Left            =   1200
      Picture         =   "Login.frx":3B4A0
      Top             =   3960
      Width           =   6000
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   1200
      Picture         =   "Login.frx":41B2B
      Top             =   2520
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   1200
      Picture         =   "Login.frx":440E8
      Top             =   840
      Width           =   2955
   End
   Begin VB.Label lbl2 
      Caption         =   "a"
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lbl1 
      Caption         =   "a"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   7080
      Width           =   2055
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

txt3.Visible = False
txt4.Visible = False
Image6.Visible = False
lbl1.Visible = False
lbl2.Visible = False

Dim n, r As String
Open App.Path & "\DATOS.txt" For Input As #1

Input #1, n, r
lbl1.Caption = n
lbl2.Caption = r
Close #1


End Sub


Private Sub Image3_Click()


If txt1.Text = lbl1.Caption Then
      If txt2.Text = lbl2.Caption Then
      Menu.Show
      Login.Hide
      Else
      MsgBox "La contraseña es incorrecta", vbInformation
      End If
Else
MsgBox "El nombre de usuario es incorrecto", vbInformation

End If


End Sub
Private Sub Image4_Click()

If MsgBox("Si creas una cuenta nueva perderas todos tus puntos y naves compradas. ¿Deseas dontinuar?", vbYesNo, "Crear cuenta nueva") = vbYes Then

txt3.Visible = True
txt4.Visible = True
Image6.Visible = True
lbl1.Visible = False
lbl2.Visible = False
Else
Exit Sub
End If
End Sub


Private Sub Image5_Click()
If MsgBox("¿Estas seguro de que quieres salir de este juego?", vbYesNo, "Salir del juego") = vbYes Then

End
End If
End Sub

Private Sub Image6_Click()

lbl1.Caption = txt3.Text
lbl2.Caption = txt4.Text
txt3.Visible = False
txt4.Visible = False
Image6.Visible = False
lbl1.Visible = False
lbl2.Visible = False

Open "./DATOS.txt" For Output As #1
Print #1, lbl1.Caption
Print #1, lbl2.Caption
Print #1, "0"
Print #1, "0"
Print #1, "0"
Print #1, "0"
Print #1, "0"
Print #1, "0"
Close #1

End Sub

Private Sub Image7_Click()
Login.Hide
Menu.Show
End Sub

Private Sub txt3_Click()
txt3.Text = ""
End Sub

Private Sub txt4_Click()
txt4.Text = ""
End Sub



