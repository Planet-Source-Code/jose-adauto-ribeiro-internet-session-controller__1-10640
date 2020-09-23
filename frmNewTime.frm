VERSION 5.00
Begin VB.Form frmNewTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configura o Tempo para Timeout"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDesistir 
      Caption         =   "Desistir"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   4920
      Width           =   2415
   End
   Begin VB.CommandButton cmdValidar 
      Caption         =   "Validar"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Frame f3dMinutos 
      Caption         =   " Minutos (múlt. de 10) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
      Begin VB.TextBox txtMinutos 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Text            =   "0"
         Top             =   360
         Width           =   1695
      End
      Begin VB.HScrollBar hscMinutos 
         Height          =   255
         LargeChange     =   10
         Left            =   360
         Max             =   50
         SmallChange     =   10
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame f3dHoras 
      Caption         =   " Número de Horas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   2415
      Begin VB.HScrollBar hscHoras 
         Height          =   255
         Left            =   360
         Max             =   12
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label lblTimeOut 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Time-out atual (hh:mm): hh:mm "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmNewTime.frx":0000
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
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmNewTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTimeout As String
Private Sub cmdDesistir_Click()
Unload Me
End Sub

Private Sub cmdValidar_Click()
Const sngMinuto As Single = 60
Const sngMilSeg As Single = 1000
If hscHoras.Value = 0 And hscMinutos.Value = 0 Then
    Unload Me
    Exit Sub
End If
frmSessionController.tmrConect.Enabled = False

intTimeOut = sngMinuto * hscHoras.Value + hscMinutos.Value
intTempoTimeOut = intTimeOut
SaveSetting App.EXEName, "Configurações", "TimeOut", intTimeOut
frmSessionController.tmrConect.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Call AcertaALinguagem
lblTimeOut.Caption = strTimeout & " (hh:mm): " & Format(CInt(intTempoTimeOut / 60), "00") + ":" + Format(intTempoTimeOut - 60 * (CInt(intTempoTimeOut / 60)), "00")
End Sub

Private Sub hscHoras_Change()
txtHoras.Text = Str(hscHoras.Value)
End Sub

Private Sub hscMinutos_Change()
txtMinutos.Text = Str(hscMinutos.Value)
End Sub
Private Sub AcertaALinguagem()
Select Case strLanguage
   Case Português
      Me.Caption = "Configuração do TimeOut"
      f3dHoras.Caption = " Número de Horas "
      f3dMinutos.Caption = " Minutos (múlt. de 10) "
      lblMensagem.Caption = " Indique abaixo daqui a quanto tempo deve-se verificar novamente o timeout de conexão com a Internet. Será somado o número de horas com os minutos indicados. Máximo: 12:50 hs. Indicar zero equivale a DESISTIR da alteração. "
      strTimeout = " Timeout atual"
      cmdValidar.Caption = "Validar"
      cmdValidar.ToolTipText = "Clique aqui para validar a alteração ..."
      cmdDesistir.Caption = "Desistir"
      cmdDesistir.ToolTipText = "Clique aqui para desistir da alteração ..."
   Case English
      Me.Caption = "TimeOut Configuration"
      f3dHoras.Caption = " Number of Hours "
      f3dMinutos.Caption = " Minutes (mult. of 10) "
      lblMensagem.Caption = " Set below how long it will be checked again if it is to finish this session. The number of hours will be added to the numbers of minutes. Maximum: 12:50hs. Zeroes will be considered to quit."
      strTimeout = " Timeout now"
      cmdValidar.Caption = "Save"
      cmdValidar.ToolTipText = "Click here to save the modifications ..."
      cmdDesistir.Caption = "Quit"
      cmdDesistir.ToolTipText = "Click here to quit the modifications  ..."
End Select
End Sub
