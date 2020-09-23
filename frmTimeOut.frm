VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTimeOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de Tempo de Conexão à Internet"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar Conectado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblTempo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tempo restante: 120 segundos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1185
      TabIndex        =   3
      Top             =   1920
      Width           =   2805
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmTimeOut.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
End
Attribute VB_Name = "frmTimeOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intSegundos As Integer
Const MaxWait As Integer = 120
Dim strTimeout As String
Dim strSegundos As String
Private Sub Command1_Click()
    blnConnected = False
    intTimeOut = intTempoTimeOut
    frmSessionController.tmrConect.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    AlteraPosiçãoDoForm Me, "Topo"
    Call AcertaALinguagem
    intSegundos = MaxWait
    Me.ProgressBar1.Max = MaxWait
End Sub

Private Sub Timer1_Timer()
    Dim lResult As Long, RetVal As Long
    Dim Tstatus As RASCONNSTATUS95
    intSegundos = intSegundos - 1
    lblTempo.Caption = strTimeout & intSegundos & strSegundos
    Me.ProgressBar1.Value = MaxWait - intSegundos
    If intSegundos = 0 Then
        frmSessionController.Timer1.Enabled = False
        blnEncerraConexão = True
        lngStatus = RasHangUp(TRasCon(0).hRasCon)
        Do
           IsConnected
           'Debug.Print Now & " - frmTimeOut.Timer1:" & Str(RetCode)
           If CLng(RetCode) = CLng(0) Then
              i = 100
           End If
           i = i + 1
        Loop Until i > 100
        blnEncerraConexão = False
        Call InsereEventoNoLog(4)
        frmSessionController.strStatus = "Off"
        frmSessionController.Timer1.Enabled = True
        intTimeOut = intTempoTimeOut
        frmSessionController.tmrConect.Enabled = True
        Unload Me
    End If
End Sub
Private Sub AcertaALinguagem()
Select Case strLanguage
   Case Português
      Me.Caption = "Expirando Tempo de Conexão à Internet"
      lblMensagem.Caption = " Será encerrada a sessão com a Internet ao expirar o tempo abaixo. Para continuar com a sessão ativa clique em ""Continuar Conectado"". "
      strTimeout = " Tempo restante: "
      strSegundos = " segundos."
      Command1.Caption = "Continuar Conectado"
   Case English
      Me.Caption = "Internet Connection Expiration Time"
      lblMensagem.Caption = " The session will be interrupted at the expiration time. Click on ""Keep the Connection"" to continue connected. "
      strTimeout = " Residual time: "
      strSegundos = " seconds."
      Command1.Caption = "Keep the Connection"
End Select
End Sub

