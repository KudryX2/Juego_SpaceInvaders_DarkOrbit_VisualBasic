VERSION 5.00
Begin VB.Form Menu 
   BorderStyle     =   0  'None
   Caption         =   "MENU "
   ClientHeight    =   8475
   ClientLeft      =   2880
   ClientTop       =   1080
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   7440
      Top             =   5760
      Width           =   6375
   End
   Begin VB.Image SALIR 
      Height          =   975
      Left            =   14160
      Top             =   240
      Width           =   735
   End
   Begin VB.Image AUTOR 
      Height          =   1335
      Left            =   7440
      Top             =   5760
      Width           =   6375
   End
   Begin VB.Image CONTROLES 
      Height          =   1215
      Left            =   7440
      Top             =   3840
      Width           =   6375
   End
   Begin VB.Image EMPEZAR 
      Height          =   1215
      Left            =   7440
      Top             =   2040
      Width           =   6375
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CONTROLES_Click()
INFO1.Show
Menu.Hide
End Sub

Private Sub EMPEZAR_Click()
Menu.Hide
JUEGO.Show

End Sub

Private Sub Form_Load()
Login.Show
Menu.Hide
End Sub


Private Sub Image1_Click()
MsgBox "Aplicacion creada por Kudryavtsev Oleksandr, Puedes contactar ocnmigo con ese email mlkzm99@gmail.com. Esta aplicacion es gratuita sin embargo queda prohibida su modificacion o venta. ", vbInformation, "APLICACION"
End Sub

Private Sub SALIR_Click()
If MsgBox("¿Estas seguro de que quieres salir del juego?", vbYesNo, "Salir") = vbYes Then

End
End If
End Sub
