VERSION 5.00
Begin VB.Form JUEGO 
   BorderStyle     =   0  'None
   Caption         =   "DARK ORBIT - MINIJUEGO"
   ClientHeight    =   10320
   ClientLeft      =   600
   ClientTop       =   150
   ClientWidth     =   19155
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "JUEGO.frx":0000
   ScaleHeight     =   10320
   ScaleWidth      =   19155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TRUCOS2 
      Interval        =   1000
      Left            =   2760
      Top             =   2040
   End
   Begin VB.TextBox txtTrucos 
      Height          =   495
      Left            =   19200
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.Timer TRUCOS 
      Interval        =   20
      Left            =   2760
      Top             =   1320
   End
   Begin VB.Timer M4 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2640
      Top             =   9600
   End
   Begin VB.Timer M3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2040
      Top             =   9600
   End
   Begin VB.Timer M2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1440
      Top             =   9600
   End
   Begin VB.Timer M1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   840
      Top             =   9600
   End
   Begin VB.Timer LM 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   240
      Top             =   9600
   End
   Begin VB.Timer COMBUSTIBLE2 
      Interval        =   40
      Left            =   4680
      Top             =   8760
   End
   Begin VB.Timer COMBUSTIBLE1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   4200
      Top             =   8760
   End
   Begin VB.Timer LC 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3720
      Top             =   8760
   End
   Begin VB.Timer RR2 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   4680
      Top             =   9720
   End
   Begin VB.Timer RR1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   4200
      Top             =   9720
   End
   Begin VB.Timer RR 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3720
      Top             =   9720
   End
   Begin VB.Timer JEFE 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   14040
      Top             =   360
   End
   Begin VB.Timer Tiempo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   1320
   End
   Begin VB.Timer LL2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   8760
   End
   Begin VB.Timer LL 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5520
      Top             =   8760
   End
   Begin VB.Timer L4 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6120
      Top             =   9720
   End
   Begin VB.Timer L3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5520
      Top             =   9720
   End
   Begin VB.Timer L2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6120
      Top             =   9240
   End
   Begin VB.Timer L1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5520
      Top             =   9240
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   14520
      Top             =   9480
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   13800
      Top             =   9480
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   13800
      Top             =   8640
   End
   Begin VB.Timer timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   17520
      Top             =   7920
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   16800
      Top             =   7920
   End
   Begin VB.Label TR1 
      Caption         =   "0"
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   2880
      Width           =   495
   End
   Begin VB.Image COMPRARSENTINEL 
      Height          =   855
      Left            =   6120
      Picture         =   "JUEGO.frx":19FD9
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Image COMPRARPET 
      Height          =   855
      Left            =   13080
      Picture         =   "JUEGO.frx":1A61A
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Image COMPRARVENGANCE 
      Height          =   855
      Left            =   10560
      Picture         =   "JUEGO.frx":1AC5B
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Image COMPRARVENOM 
      Height          =   855
      Left            =   8280
      Picture         =   "JUEGO.frx":1B2A0
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Image COMPRARSPECTRUM 
      Height          =   855
      Left            =   6120
      Picture         =   "JUEGO.frx":1B8D2
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Image SENTINEL 
      Height          =   1695
      Left            =   6120
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Image PET 
      Height          =   1695
      Left            =   12960
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image VENGANCE 
      Height          =   1695
      Left            =   10440
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Image VENOM 
      Height          =   1815
      Left            =   8160
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Image SPECTRUM 
      Height          =   1695
      Left            =   6120
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image GOLIATH 
      Height          =   1815
      Left            =   4320
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Image SN 
      Height          =   6480
      Left            =   3840
      Picture         =   "JUEGO.frx":1BF04
      Top             =   1560
      Width           =   11520
   End
   Begin VB.Label nave5 
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   8160
      Width           =   975
   End
   Begin VB.Label nave4 
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label nave3 
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   7200
      Width           =   975
   End
   Begin VB.Label nave2 
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label nave1 
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label CONTRASEÑA 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Label NOMBRE 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Image MINA4 
      Height          =   570
      Left            =   9360
      Picture         =   "JUEGO.frx":66BD7
      Top             =   120
      Width           =   675
   End
   Begin VB.Image MINA3 
      Height          =   570
      Left            =   6000
      Picture         =   "JUEGO.frx":67815
      Top             =   120
      Width           =   675
   End
   Begin VB.Image MINA2 
      Height          =   570
      Left            =   2640
      Picture         =   "JUEGO.frx":68453
      Top             =   120
      Width           =   675
   End
   Begin VB.Image MINA1 
      Height          =   570
      Left            =   12960
      Picture         =   "JUEGO.frx":69091
      Top             =   120
      Width           =   675
   End
   Begin VB.Image FU 
      Height          =   1080
      Left            =   8880
      Picture         =   "JUEGO.frx":69CCF
      Top             =   1080
      Width           =   990
   End
   Begin VB.Image REPARACION 
      Height          =   1455
      Left            =   6600
      Picture         =   "JUEGO.frx":6ABD0
      Top             =   -120
      Width           =   1050
   End
   Begin VB.Image HPJ4 
      Height          =   285
      Left            =   10680
      Picture         =   "JUEGO.frx":6BEA0
      Top             =   120
      Width           =   825
   End
   Begin VB.Image HPJ3 
      Height          =   285
      Left            =   9840
      Picture         =   "JUEGO.frx":6C23F
      Top             =   120
      Width           =   825
   End
   Begin VB.Image HPJ2 
      Height          =   285
      Left            =   9000
      Picture         =   "JUEGO.frx":6C5DE
      Top             =   120
      Width           =   825
   End
   Begin VB.Image HPJ1 
      Height          =   285
      Left            =   8160
      Picture         =   "JUEGO.frx":6C97D
      Top             =   120
      Width           =   825
   End
   Begin VB.Label HPJ 
      Caption         =   "20"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Image CUBI 
      Height          =   1095
      Left            =   7560
      Picture         =   "JUEGO.frx":6CD1C
      Top             =   -4200
      Width           =   1215
   End
   Begin VB.Label FUEL 
      Caption         =   "4"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.Image FUEL4 
      Height          =   570
      Left            =   5040
      Picture         =   "JUEGO.frx":6E093
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image FUEL3 
      Height          =   570
      Left            =   3360
      Picture         =   "JUEGO.frx":6F61C
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image FUEL2 
      Height          =   570
      Left            =   1680
      Picture         =   "JUEGO.frx":70BA5
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image FUEL1 
      Height          =   570
      Left            =   0
      Picture         =   "JUEGO.frx":7212E
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label T 
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label AE 
      Caption         =   "0"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Puntos 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Image LASER2 
      Height          =   1005
      Left            =   10080
      Picture         =   "JUEGO.frx":742CD
      Top             =   7800
      Width           =   105
   End
   Begin VB.Image LASER1 
      Height          =   1005
      Left            =   9120
      Picture         =   "JUEGO.frx":75446
      Top             =   7800
      Width           =   105
   End
   Begin VB.Image HP5 
      Height          =   555
      Left            =   6240
      Picture         =   "JUEGO.frx":765BF
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label HP 
      Caption         =   "5"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Image HP4 
      Height          =   555
      Left            =   4680
      Picture         =   "JUEGO.frx":76AE3
      Top             =   0
      Width           =   1545
   End
   Begin VB.Image HP3 
      Height          =   555
      Left            =   3120
      Picture         =   "JUEGO.frx":77007
      Top             =   0
      Width           =   1545
   End
   Begin VB.Image HP2 
      Height          =   555
      Left            =   1560
      Picture         =   "JUEGO.frx":7752B
      Top             =   0
      Width           =   1545
   End
   Begin VB.Image HP1 
      Height          =   555
      Left            =   0
      Picture         =   "JUEGO.frx":77A4F
      Top             =   0
      Width           =   1545
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   14
      Left            =   10320
      Picture         =   "JUEGO.frx":79695
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   13
      Left            =   4080
      Picture         =   "JUEGO.frx":7A35A
      Top             =   960
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   12
      Left            =   4680
      Picture         =   "JUEGO.frx":7B01F
      Top             =   2040
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   11
      Left            =   5160
      Picture         =   "JUEGO.frx":7BCE4
      Top             =   360
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   10
      Left            =   6240
      Picture         =   "JUEGO.frx":7C9A9
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   9
      Left            =   11400
      Picture         =   "JUEGO.frx":7D66E
      Top             =   2280
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   8
      Left            =   12720
      Picture         =   "JUEGO.frx":7E333
      Top             =   1680
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   7
      Left            =   9600
      Picture         =   "JUEGO.frx":7EFF8
      Top             =   600
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   6
      Left            =   7320
      Picture         =   "JUEGO.frx":7FCBD
      Top             =   840
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   5
      Left            =   11880
      Picture         =   "JUEGO.frx":80982
      Top             =   720
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   2
      Left            =   9000
      Picture         =   "JUEGO.frx":81647
      Top             =   1800
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   3
      Left            =   10800
      Picture         =   "JUEGO.frx":8230C
      Top             =   720
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   4
      Left            =   7680
      Picture         =   "JUEGO.frx":82FD1
      Top             =   2280
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   1
      Left            =   8520
      Picture         =   "JUEGO.frx":83C96
      Top             =   480
      Width           =   750
   End
   Begin VB.Image c 
      Height          =   690
      Index           =   0
      Left            =   6240
      Picture         =   "JUEGO.frx":8495B
      Top             =   720
      Width           =   750
   End
   Begin VB.Image START 
      Height          =   1665
      Left            =   6720
      Picture         =   "JUEGO.frx":85620
      Top             =   720
      Width           =   6120
   End
   Begin VB.Image NAVE 
      Height          =   2160
      Left            =   8520
      Picture         =   "JUEGO.frx":93DCD
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   2340
   End
End
Attribute VB_Name = "JUEGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()


START.Visible = False

NOMBRE.Visible = False
CONTRASEÑA.Visible = False


HP.Visible = False
FUEL.Visible = False
HPJ.Visible = False

REPARACION.Visible = False
FU.Visible = False

LASER1.Visible = False
LASER2.Visible = False
NAVE.Visible = False

CUBI.Visible = False


Dim k As Integer

HPJ1.Visible = False
HPJ2.Visible = False
HPJ3.Visible = False
HPJ4.Visible = False

For k = k To 14
c(k).Visible = False
Next

MINA1.Visible = False
MINA2.Visible = False
MINA3.Visible = False
MINA4.Visible = False



Dim n, r, p, n1, n2, n3, n4, n5 As String
Open App.Path & "\DATOS.txt" For Input As #1

Input #1, n, r, p, n1, n2, n3, n4, n5
NOMBRE.Caption = n
CONTRASEÑA.Caption = r
Puntos.Caption = p
nave1.Caption = n1
nave2.Caption = n2
nave3.Caption = n3
nave4.Caption = n4
nave5.Caption = n5
Close #1

If nave1.Caption = 0 Then
SPECTRUM.Visible = False
COMPRARSPECTRUM.Visible = True
Else
SPECTRUM.Visible = True
COMPRARSPECTRUM.Visible = False
End If

If nave2.Caption = 0 Then
VENOM.Visible = False
COMPRARVENOM.Visible = True
Else
VENOM.Visible = True
COMPRARVENOM.Visible = False
End If

If nave3.Caption = 0 Then
VENGANCE.Visible = False
COMPRARVENGANCE.Visible = True
Else
VENGANCE.Visible = True
COMPRARVENGANCE.Visible = False
End If

If nave4.Caption = 0 Then
PET.Visible = False
COMPRARPET.Visible = True
Else
PET.Visible = True
COMPRARPET.Visible = False
End If

If nave5.Caption = 0 Then
SENTINEL.Visible = False
COMPRARSENTINEL.Visible = True
Else
PET.Visible = True
COMPRARSENTINEL.Visible = False
End If

TR1.Visible = False

nave1.Visible = False
nave2.Visible = False
nave3.Visible = False
nave4.Visible = False
nave5.Visible = False

AE.Visible = False
T.Visible = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case vbKeyRight
   If NAVE.Left < 12000 Then
   Timer1.Enabled = True
   Else
   Timer1.Enabled = False
   End If
   
   If LASER1.Left < 12500 Then
   L2.Enabled = True
   Else
   L2.Enabled = False
   End If
   
   If LASER2.Left < 13500 Then
   L4.Enabled = True
   Else
   L4.Enabled = False
   End If
   
   
   
Case vbKeyLeft
   If NAVE.Left > 2000 Then
   timer2.Enabled = True
   Else
   timer2.Enabled = False
   End If
   
   If LASER1.Left > 2800 Then
   L1.Enabled = True
   Else
   L1.Enabled = False
   End If
   
   If LASER2.Left > 3500 Then
   L3.Enabled = True
   Else
   L3.Enabled = False
   End If
   
Case vbKeySpace
   LASER1.Visible = True
   LASER2.Visible = True
   LL.Enabled = True
   LL2.Enabled = True
    
End Select
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode

Case vbKeyRight
Timer1.Enabled = False
L2.Enabled = False
L4.Enabled = False

Case vbKeyLeft
timer2.Enabled = False
L1.Enabled = False
L3.Enabled = False

End Select
End Sub



Private Sub JEFE_Timer()

CUBI.Top = CUBI.Top + 14

If HPJ.Caption < 15 Then
HPJ4.Visible = False
End If

If HPJ.Caption < 10 Then
HPJ3.Visible = False
End If

If HPJ.Caption < 5 Then
HPJ2.Visible = False
End If

If HPJ.Caption < 0 Then
HPJ1.Visible = False
End If


If CUBI.Top > 10000 Then
HP.Caption = 0
End If


If HPJ.Caption <= 0 Then
JEFE.Enabled = False
AE.Caption = AE.Caption + 1
Puntos.Caption = Puntos.Caption + 250
CUBI.Top = 20
CUBI.Visible = False
HPJ1.Visible = False
HPJ2.Visible = False
HPJ3.Visible = False
HPJ4.Visible = False
Timer3.Interval = 40
HP.Caption = HP.Caption + 1
End If

End Sub


Private Sub L1_Timer()
LASER1.Left = LASER1.Left - 300
End Sub

Private Sub L2_Timer()
LASER1.Left = LASER1.Left + 300
End Sub

Private Sub L3_Timer()
LASER2.Left = LASER2.Left - 300
End Sub

Private Sub L4_Timer()
LASER2.Left = LASER2.Left + 300
End Sub



Private Sub LL_Timer()


LASER1.Top = LASER1.Top - 350

If LASER1.Top < 200 Then
LASER1.Visible = False
LASER1.Top = 7820
LL.Enabled = False
End If


End Sub

Private Sub LL2_Timer()
LASER2.Top = LASER2.Top - 350

If LASER2.Top < 200 Then
LASER2.Visible = False
LASER2.Top = 7820
LL.Enabled = False
End If


End Sub
Private Sub LC_Timer()
FU.Top = FU.Top + 75

If FU.Top > 10000 Then
FU.Top = -120
FU.Visible = False
COMBUSTIBLE1.Enabled = False
COMBUSTIBLE2.Enabled = False
LC.Enabled = False
End If

If FU.Top > 7500 And FU.Left > NAVE.Left And FU.Left < NAVE.Left + 2000 Then
FUEL.Caption = FUEL.Caption + 1
COMBUSTIBLE1.Enabled = False
COMBUSTIBLE2.Enabled = False
Puntos.Caption = Puntos.Caption + 100
FU.Visible = False
FU.Top = -120
LC.Enabled = False
End If

End Sub
Private Sub COMBUSTIBLE1_Timer()
FU.Left = FU.Left - 50

If FU.Left < 2000 Then
COMBUSTIBLE1.Enabled = False
COMBUSTIBLE2.Enabled = True
End If
End Sub

Private Sub COMBUSTIBLE2_Timer()
FU.Left = FU.Left + 50

If FU.Left > 12000 Then
COMBUSTIBLE2.Enabled = False
COMBUSTIBLE1.Enabled = True
End If
End Sub



Private Sub RR_Timer()
REPARACION.Top = REPARACION.Top + 45

If REPARACION.Top > 7500 And REPARACION.Left > NAVE.Left And REPARACION.Left < NAVE.Left + 2000 Then
HP.Caption = HP.Caption + 1
RR1.Enabled = False
RR2.Enabled = False
REPARACION.Visible = False
REPARACION.Top = -120
RR.Enabled = False
End If

If REPARACION.Top > 12000 Then
REPARACION.Top = -120
REPARACION.Visible = False
RR1.Enabled = False
RR2.Enabled = False
RR.Enabled = False
End If


End Sub

Private Sub RR1_Timer()
REPARACION.Left = REPARACION.Left - 50

If REPARACION.Left < 2000 Then
RR2.Enabled = True
RR1.Enabled = False
End If

End Sub

Private Sub RR2_Timer()
REPARACION.Left = REPARACION.Left + 50

If REPARACION.Left > 12000 Then
RR1.Enabled = True
RR2.Enabled = False
End If

End Sub




Private Sub timer2_Timer()
NAVE.Left = NAVE.Left - 300

End Sub


Private Sub Timer1_Timer()
NAVE.Left = NAVE.Left + 300

End Sub

Private Sub PAUSA_Click()

PAUSA.Visible = False
START.Visible = True
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False

JEFE.Enabled = False

Tiempo.Enabled = False

LM.Enabled = False
L1.Enabled = False
L2.Enabled = False
L3.Enabled = False
L4.Enabled = False
End Sub

Private Sub START_Click()




START.Visible = False

NAVE.Visible = True

Timer3.Enabled = True
Timer4.Enabled = True

For k = k To 14
c(k).Visible = True
Next

Tiempo.Enabled = True

LM.Enabled = True

End Sub


Private Sub Timer3_Timer()


For k = k To 14
c(k).Top = c(k).Top + 25
Next



If c(0).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(0).Top = c(0).Top - 8000
End If

If c(1).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(1).Top = c(1).Top - 8000
End If

If c(2).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(2).Top = c(2).Top - 8000
End If

If c(3).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(3).Top = c(3).Top - 8000
End If


If c(4).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(4).Top = c(4).Top - 8000
End If

If c(5).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(5).Top = c(5).Top - 8000
End If

If c(6).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(6).Top = c(6).Top - 8000
End If

If c(7).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(7).Top = c(7).Top - 8000
End If

If c(8).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(8).Top = c(8).Top - 8000
End If

If c(9).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(9).Top = c(9).Top - 8000
End If

If c(10).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(10).Top = c(10).Top - 8000
End If

If c(11).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(11).Top = c(11).Top - 8000
End If

If c(12).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(12).Top = c(12).Top - 8000
End If

If c(13).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(13).Top = c(13).Top - 8000
End If

If c(14).Top > 7500 Then
HP.Caption = HP.Caption - 1
c(14).Top = c(14).Top - 8000
End If


If HP.Caption = 5 Then
HP1.Visible = True
HP2.Visible = True
HP3.Visible = True
HP4.Visible = True
HP5.Visible = True
End If

If HP.Caption = 4 Then
HP1.Visible = True
HP2.Visible = True
HP3.Visible = True
HP4.Visible = True
HP5.Visible = False
End If
 
If HP.Caption = 3 Then
HP1.Visible = True
HP2.Visible = True
HP3.Visible = True
HP4.Visible = False
HP5.Visible = False
End If
 
If HP.Caption = 2 Then
HP1.Visible = True
HP2.Visible = True
HP3.Visible = False
HP4.Visible = False
HP5.Visible = False
End If
 
If HP.Caption = 1 Then
HP1.Visible = True
HP2.Visible = False
HP3.Visible = False
HP4.Visible = False
HP5.Visible = False
End If


If HP.Caption = 0 Then
MsgBox "PERDISTE - TU NAVE FUE DERRIBADA", vbInformation
End
End If
 


If LASER1.Top < c(0).Top And LASER1.Left < c(0).Left + 1000 And LASER2.Left > c(0).Left Then
c(0).Top = 100
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(1).Top And LASER1.Left < c(1).Left + 1000 And LASER1.Left > c(1).Left Then
c(1).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(2).Top And LASER1.Left < c(2).Left + 1000 And LASER1.Left > c(2).Left Then
c(2).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(3).Top And LASER1.Left < c(3).Left + 1000 And LASER1.Left > c(3).Left Then
c(3).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(4).Top And LASER1.Left < c(4).Left + 1000 And LASER1.Left > c(4).Left Then
c(4).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(5).Top And LASER1.Left < c(5).Left + 1000 And LASER2.Left > c(5).Left Then
c(5).Top = 100
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(6).Top And LASER1.Left < c(6).Left + 1000 And LASER1.Left > c(6).Left Then
c(6).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(7).Top And LASER1.Left < c(7).Left + 1000 And LASER1.Left > c(7).Left Then
c(7).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(8).Top And LASER1.Left < c(8).Left + 1000 And LASER1.Left > c(8).Left Then
c(8).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(8).Top And LASER1.Left < c(8).Left + 1000 And LASER1.Left > c(8).Left Then
c(8).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(9).Top And LASER1.Left < c(9).Left + 1000 And LASER2.Left > c(9).Left Then
c(9).Top = 100
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(10).Top And LASER1.Left < c(10).Left + 1000 And LASER1.Left > c(10).Left Then
c(10).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(11).Top And LASER1.Left < c(11).Left + 1000 And LASER1.Left > c(11).Left Then
c(11).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(12).Top And LASER1.Left < c(12).Left + 1000 And LASER1.Left > c(12).Left Then
c(12).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(13).Top And LASER1.Left < c(13).Left + 1000 And LASER1.Left > c(13).Left Then
c(13).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER1.Top < c(14).Top And LASER1.Left < c(14).Left + 1000 And LASER1.Left > c(14).Left Then
c(14).Top = 50
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If





If LASER2.Top < c(0).Top And LASER2.Left < c(0).Left + 1000 And LASER2.Left > c(0).Left Then
c(0).Top = 100
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(1).Top And LASER2.Left < c(1).Left + 1000 And LASER2.Left > c(1).Left Then
c(1).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(2).Top And LASER2.Left < c(2).Left + 1000 And LASER2.Left > c(2).Left Then
c(2).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(3).Top And LASER2.Left < c(3).Left + 1000 And LASER2.Left > c(3).Left Then
c(3).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(4).Top And LASER2.Left < c(4).Left + 1000 And LASER2.Left > c(4).Left Then
c(4).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(5).Top And LASER2.Left < c(5).Left + 1000 And LASER2.Left > c(5).Left Then
c(5).Top = 100
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(6).Top And LASER2.Left < c(6).Left + 1000 And LASER2.Left > c(6).Left Then
c(6).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(7).Top And LASER2.Left < c(7).Left + 1000 And LASER2.Left > c(7).Left Then
c(7).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(8).Top And LASER2.Left < c(8).Left + 1000 And LASER2.Left > c(8).Left Then
c(8).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(8).Top And LASER2.Left < c(8).Left + 1000 And LASER2.Left > c(8).Left Then
c(8).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(9).Top And LASER2.Left < c(9).Left + 1000 And LASER2.Left > c(9).Left Then
c(9).Top = 100
LASER1.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(10).Top And LASER2.Left < c(10).Left + 1000 And LASER2.Left > c(10).Left Then
c(10).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(11).Top And LASER2.Left < c(11).Left + 1000 And LASER2.Left > c(11).Left Then
c(11).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(12).Top And LASER2.Left < c(12).Left + 1000 And LASER2.Left > c(12).Left Then
c(12).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(13).Top And LASER2.Left < c(13).Left + 1000 And LASER2.Left > c(13).Left Then
c(13).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If

If LASER2.Top < c(14).Top And LASER2.Left < c(14).Left + 1000 And LASER2.Left > c(14).Left Then
c(14).Top = 50
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
Puntos.Caption = Puntos.Caption + 10
AE.Caption = AE.Caption + 1
End If




If LASER1.Top < CUBI.Top And LASER1.Left < CUBI.Left + 1500 And LASER1.Left > CUBI.Left Then
HPJ.Caption = HPJ.Caption - 1
LASER1.Top = 7800
LL.Enabled = False
LASER1.Visible = False
End If

If LASER2.Top < CUBI.Top And LASER2.Left < CUBI.Left + 1500 And LASER2.Left > CUBI.Left Then
HPJ.Caption = HPJ.Caption - 1
LASER2.Top = 7800
LL2.Enabled = False
LASER2.Visible = False
End If




If HPJ.Caption < 0 Then
CUBI.Visible = False
CUBI.Top = -4200
Timer3.Interval = 40
End If


LASER1.Left = NAVE.Left + 600

LASER2.Left = NAVE.Left + 1500



 
End Sub

Private Sub Timer4_Timer()

For k = k To 14
c(k).Left = c(k).Left + 120
Next

If c(8).Left > 14000 Then
Timer4.Enabled = False
Timer5.Enabled = True
End If

End Sub


Private Sub Timer5_Timer()

For k = k To 14
c(k).Left = c(k).Left - 120
Next

If c(13).Left < 2000 Then
Timer5.Enabled = False
Timer4.Enabled = True
End If
End Sub



Private Sub LM_Timer()
Randomize (n)
n = Int(Rnd * 4) + 1

If n = 1 Then
M1.Enabled = True
MINA1.Visible = True
End If

If n = 2 Then
M2.Enabled = True
MINA2.Visible = True
End If

If n = 3 Then
M3.Enabled = True
MINA3.Visible = True
End If

If n = 4 Then
M4.Enabled = True
MINA4.Visible = True
End If

End Sub


Private Sub M1_Timer()
MINA1.Top = MINA1.Top + 15
If MINA1.Top > 10000 Then
MINA1.Top = 30
End If
If MINA1.Top > NAVE.Top And MINA1.Left > NAVE.Left And MINA1.Left < NAVE.Left + 2000 Then
HP.Caption = HP.Caption - 2
 If HP.Caption < 0 Then
 HP.Caption = 0
 End If
MINA1.Top = 30
End If
End Sub

Private Sub M2_Timer()
MINA2.Top = MINA2.Top + 15
If MINA2.Top > 10000 Then
MINA2.Top = 30
End If
If MINA2.Top > NAVE.Top And MINA2.Left > NAVE.Left And MINA2.Left < NAVE.Left + 2000 Then
HP.Caption = HP.Caption - 2
 If HP.Caption < 0 Then
 HP.Caption = 0
 End If
MINA2.Top = 30
End If
End Sub

Private Sub M3_Timer()
MINA3.Top = MINA3.Top + 15
If MINA3.Top > 10000 Then
MINA3.Top = 30
End If
If MINA3.Top > NAVE.Top And MINA3.Left > NAVE.Left And MINA3.Left < NAVE.Left + 2000 Then
HP.Caption = HP.Caption - 2
 If HP.Caption < 0 Then
 HP.Caption = 0
 End If
MINA3.Top = 30
End If
End Sub

Private Sub M4_Timer()
MINA4.Top = MINA4.Top + 15
If MINA4.Top > 10000 Then
MINA4.Top = 30
End If
If MINA4.Top > NAVE.Top And MINA4.Left > NAVE.Left And MINA4.Left < NAVE.Left + 2000 Then
HP.Caption = HP.Caption - 2
 If HP.Caption < 0 Then
 HP.Caption = 0
 End If
MINA4.Top = 30
End If
End Sub




Private Sub SPECTRUM_Click()
START.Visible = True
SN.Visible = False

NAVE.Picture = LoadPicture(App.Path & "\2.gif")

COMPRARSPECTRUM.Visible = False
COMPRARVENOM.Visible = False
COMPRARVENGANCE.Visible = False
COMPRARPET.Visible = False
COMPRARSENTINEL.Visible = False
End Sub


Private Sub GOLIATH_Click()
START.Visible = True
SN.Visible = False

NAVE.Picture = LoadPicture(App.Path & "\1.gif")

COMPRARSPECTRUM.Visible = False
COMPRARVENOM.Visible = False
COMPRARVENGANCE.Visible = False
COMPRARPET.Visible = False
COMPRARSENTINEL.Visible = False
End Sub




Private Sub VENOM_Click()
START.Visible = True
SN.Visible = False

NAVE.Picture = LoadPicture(App.Path & "\3.gif")

COMPRARSPECTRUM.Visible = False
COMPRARVENOM.Visible = False
COMPRARVENGANCE.Visible = False
COMPRARPET.Visible = False
COMPRARSENTINEL.Visible = False
End Sub


Private Sub VENGANCE_Click()
START.Visible = True
SN.Visible = False

NAVE.Picture = LoadPicture(App.Path & "\4.gif")

COMPRARSPECTRUM.Visible = False
COMPRARVENOM.Visible = False
COMPRARVENGANCE.Visible = False
COMPRARPET.Visible = False
COMPRARSENTINEL.Visible = False
End Sub

Private Sub PET_Click()
START.Visible = True
SN.Visible = False

NAVE.Picture = LoadPicture(App.Path & "\6.gif")

COMPRARSPECTRUM.Visible = False
COMPRARVENOM.Visible = False
COMPRARVENGANCE.Visible = False
COMPRARPET.Visible = False
COMPRARSENTINEL.Visible = False
End Sub


Private Sub SENTINEL_Click()
START.Visible = True
SN.Visible = False

NAVE.Picture = LoadPicture(App.Path & "\5.gif")

COMPRARSPECTRUM.Visible = False
COMPRARVENOM.Visible = False
COMPRARVENGANCE.Visible = False
COMPRARPET.Visible = False
COMPRARSENTINEL.Visible = False
End Sub





Private Sub COMPRARSPECTRUM_Click()
If Puntos.Caption >= 5000 Then

Puntos.Caption = Puntos.Caption - 5000

Open "./DATOS.txt" For Output As #1
Print #1, NOMBRE.Caption
Print #1, CONTRASEÑA.Caption
Print #1, Puntos.Caption
Print #1, 1
Print #1, nave2.Caption
Print #1, nave3.Caption
Print #1, nave4.Caption
Print #1, nave5.Caption
Close #1

Call Form_Load

Else
MsgBox "No tienes puntos suficientes", vbInformation
End If
End Sub


Private Sub COMPRARVENOM_Click()
If Puntos.Caption >= 5000 Then

Puntos.Caption = Puntos.Caption - 5000

Open "./DATOS.txt" For Output As #1
Print #1, NOMBRE.Caption
Print #1, CONTRASEÑA.Caption
Print #1, Puntos.Caption
Print #1, nave1.Caption
Print #1, 1
Print #1, nave3.Caption
Print #1, nave4.Caption
Print #1, nave5.Caption
Close #1

Call Form_Load

Else
MsgBox "No tienes puntos suficientes", vbInformation
End If
End Sub


Private Sub COMPRARVENGANCE_Click()
If Puntos.Caption >= 10000 Then

Puntos.Caption = Puntos.Caption - 10000

Open "./DATOS.txt" For Output As #1
Print #1, NOMBRE.Caption
Print #1, CONTRASEÑA.Caption
Print #1, Puntos.Caption
Print #1, nave1.Caption
Print #1, nave2.Caption
Print #1, 1
Print #1, nave4.Caption
Print #1, nave5.Caption
Close #1

Call Form_Load

Else
MsgBox "No tienes puntos suficientes", vbInformation
End If
End Sub


Private Sub COMPRARPET_Click()
If Puntos.Caption >= 15000 Then

Puntos.Caption = Puntos.Caption - 15000

Open "./DATOS.txt" For Output As #1
Print #1, NOMBRE.Caption
Print #1, CONTRASEÑA.Caption
Print #1, Puntos.Caption
Print #1, nave1.Caption
Print #1, nave2.Caption
Print #1, nave3.Caption
Print #1, 1
Print #1, nave5.Caption
Close #1

Call Form_Load

Else
MsgBox "No tienes puntos suficientes", vbInformation
End If
End Sub



Private Sub COMPRARSENTINEL_Click()
If Puntos.Caption >= 15000 Then

Puntos.Caption = Puntos.Caption - 15000

Open "./DATOS.txt" For Output As #1
Print #1, NOMBRE.Caption
Print #1, CONTRASEÑA.Caption
Print #1, Puntos.Caption
Print #1, nave1.Caption
Print #1, nave2.Caption
Print #1, nave3.Caption
Print #1, nave4.Caption
Print #1, 1
Close #1

Call Form_Load

Else
MsgBox "No tienes puntos suficientes", vbInformation
End If
End Sub




Private Sub Tiempo_Timer()

T.Caption = T.Caption + 1

Open "./DATOS.txt" For Output As #1
Print #1, NOMBRE.Caption
Print #1, CONTRASEÑA.Caption
Print #1, Puntos.Caption
Print #1, nave1.Caption
Print #1, nave2.Caption
Print #1, nave3.Caption
Print #1, nave4.Caption
Print #1, nave5.Caption
Close #1


If T.Caption = 95 Then
T.Caption = 10
LC.Enabled = True
COMBUSTIBLE1.Enabled = True
FU.Visible = True
End If

If T.Caption = 10 Then
FUEL.Caption = FUEL.Caption - 1
Timer3.Interval = 34
End If


If T.Caption = 40 Then
HPJ.Caption = 20
CUBI.Picture = LoadPicture(App.Path & "\CUBI1.gif")
CUBI.Visible = True
JEFE.Enabled = True
Timer3.Interval = 40
CUBI.Top = 50
HPJ1.Visible = True
HPJ2.Visible = True
HPJ3.Visible = True
HPJ4.Visible = True

LC.Enabled = True
COMBUSTIBLE1.Enabled = True
FU.Visible = True

For k = k To 14
c(k).Picture = LoadPicture(App.Path & "\Alien1.gif")
c(k).Top = c(k).Top - 3000
Next
End If


If T.Caption = 50 Then
FUEL.Caption = FUEL.Caption - 1
End If



If T.Caption = 60 Then
HPJ.Caption = 20
CUBI.Picture = LoadPicture(App.Path & "\CUBI2.gif")
CUBI.Visible = True
JEFE.Enabled = True
Timer3.Interval = 45
CUBI.Top = 25
HPJ1.Visible = True
HPJ2.Visible = True
HPJ3.Visible = True
HPJ4.Visible = True


For k = k To 14
c(k).Picture = LoadPicture(App.Path & "\Alien2.gif")
c(k).Top = c(k).Top - 3000
Next

End If


If T.Caption = 70 Then
LC.Enabled = True
COMBUSTIBLE1.Enabled = True
FU.Visible = True
FUEL.Caption = FUEL.Caption - 1

REPARACION.Visible = True
RR.Enabled = True
RR1.Enabled = True
End If


If T.Caption = 80 Then
HPJ.Caption = 20
CUBI.Picture = LoadPicture(App.Path & "\CUBI3.gif")
CUBI.Visible = True
JEFE.Enabled = True
Timer3.Interval = 45
CUBI.Top = 25
HPJ1.Visible = True
HPJ2.Visible = True
HPJ3.Visible = True
HPJ4.Visible = True


For k = k To 14
c(k).Picture = LoadPicture(App.Path & "\Alien3.gif")
c(k).Top = c(k).Top - 3000
Next
End If




If FUEL.Caption = 4 Then
FUEL1.Visible = True
FUEL2.Visible = True
FUEL3.Visible = True
FUEL4.Visible = True
End If

If FUEL.Caption = 3 Then
FUEL1.Visible = True
FUEL2.Visible = True
FUEL3.Visible = True
FUEL4.Visible = False
End If

If FUEL.Caption = 2 Then
FUEL1.Visible = True
FUEL2.Visible = True
FUEL3.Visible = False
FUEL4.Visible = False
End If

If FUEL.Caption = 1 Then
FUEL1.Visible = True
FUEL2.Visible = False
FUEL3.Visible = False
FUEL4.Visible = False
End If

If FUEL.Caption = 0 Then
MsgBox "PERDISTE - TU NAVE SE QUEDO SIN COMBUSTIBLE", vbInformation
End
End If

End Sub



Private Sub TRUCOS_Timer()

If txtTrucos.Text = "vida" Then
HP.Caption = 5
txtTrucos.Text = ""
End If

If txtTrucos.Text = "fuel" Then
FUEL.Caption = 4
txtTrucos.Text = ""
End If

If txtTrucos.Text = "puntos" Then
Puntos.Caption = Puntos.Caption + 1000
txtTrucos.Text = ""
End If

If txtTrucos.Text = "jefe1" Then
T.Caption = 39
txtTrucos.Text = ""
End If

If txtTrucos.Text = "jefe2" Then
T.Caption = 59
txtTrucos.Text = ""
End If

If txtTrucos.Text = "jefe3" Then
T.Caption = 79
txtTrucos.Text = ""
End If


If Not txtTrucos.Text = "" Then
TRUCOS2.Enabled = True
End If



End Sub



Private Sub TRUCOS2_Timer()
TR1.Caption = TR1.Caption + 1
If TR1.Caption = 4 Then
TR1.Caption = 0
TRUCOS2.Enabled = False
txtTrucos.Text = ""
End If
End Sub
