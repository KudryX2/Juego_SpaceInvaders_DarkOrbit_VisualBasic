VERSION 5.00
Begin VB.Form frmfichero 
   Caption         =   "record"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin VB.Label s 
      Caption         =   "Label5"
      Height          =   375
      Index           =   4
      Left            =   5880
      TabIndex        =   14
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label s 
      Caption         =   "Label4"
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label s 
      Caption         =   "Label3"
      Height          =   375
      Index           =   2
      Left            =   6000
      TabIndex        =   12
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label s 
      Caption         =   "Label2"
      Height          =   495
      Index           =   1
      Left            =   6240
      TabIndex        =   11
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label s 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   6240
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label p 
      Caption         =   "Label2"
      Height          =   615
      Index           =   4
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
   End
   Begin VB.Label p 
      Caption         =   "Label2"
      Height          =   615
      Index           =   3
      Left            =   3240
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label p 
      Caption         =   "Label2"
      Height          =   615
      Index           =   2
      Left            =   3240
      TabIndex        =   7
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label p 
      Caption         =   "Label2"
      Height          =   615
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label n 
      Caption         =   "Label1"
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label n 
      Caption         =   "Label1"
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label n 
      Caption         =   "Label1"
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label n 
      Caption         =   "Label1"
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label p 
      Caption         =   "Label2"
      Height          =   615
      Index           =   0
      Left            =   3360
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label n 
      Caption         =   "Label1"
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmfichero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim i As Integer
Dim a, b, c As String
Open App.Path & "\pass.txt" For Input As #1

For i = 0 To 4
Input #1, a, b
n(i).Caption = a
p(i).Caption = b
s(i).Caption = c
Next i

Close #1
End Sub




