VERSION 5.00
Begin VB.Form INFO1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8070
   ClientLeft      =   2430
   ClientTop       =   990
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   Picture         =   "INFO1.frx":0000
   ScaleHeight     =   8070
   ScaleWidth      =   15000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image SP 
      Height          =   615
      Left            =   9720
      Top             =   7440
      Width           =   5655
   End
   Begin VB.Image VM2 
      Height          =   615
      Left            =   4800
      Top             =   7440
      Width           =   5295
   End
   Begin VB.Image VM 
      Height          =   855
      Left            =   240
      Top             =   7440
      Width           =   6735
   End
End
Attribute VB_Name = "INFO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
VM2.Visible = False
VM.Visible = True
SP.Visible = True
End Sub

Private Sub SP_Click()
INFO1.Picture = LoadPicture(App.Path & "\i2.gif")
SP.Visible = False
VM.Visible = False
VM2.Visible = True
End Sub

Private Sub VM_Click()
INFO1.Picture = LoadPicture(App.Path & "\i1.gif")
INFO1.Hide
Menu.Show
End Sub

Private Sub VM2_Click()
INFO1.Picture = LoadPicture(App.Path & "\i1.gif")
INFO1.Hide
SP.Visible = True
Menu.Show
End Sub
