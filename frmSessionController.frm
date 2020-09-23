VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSessionController 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Session Controller"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
   Icon            =   "frmSessionController.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlListIcons 
      Left            =   3720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSessionController.frx":0442
            Key             =   "NewTimeOut"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSessionController.frx":0896
            Key             =   "ResumoMensal"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSessionController.frx":0CEA
            Key             =   "Esconder"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSessionController.frx":113E
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSessionController.frx":145A
            Key             =   "Lingua"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbFunções 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlListIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NewTimeOut"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ResumoMensal"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Esconder"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Lingua"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrConect 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   1320
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label lblTempoTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tempo total conectado: hh:mm:ss "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   5040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblTimeOut 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   " Tempo restante: nn minutos "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Label lblTempoConexão 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "  Tempo de conexão: hh:mm:ss  "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   3150
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Situação atual:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1560
   End
   Begin VB.Menu PopMenu01 
      Caption         =   "PopMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuTelaInicial 
         Caption         =   "&Mostrar Tela Inicial ..."
      End
      Begin VB.Menu mnuEncerrar 
         Caption         =   "Ence&rrar"
      End
   End
End
Attribute VB_Name = "frmSessionController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
      Public strStatus As String
      Dim strMsgFinal As String
      Dim strMsgConfirmação As String
      Dim strMsgLinguagem As String
      Dim dtaTempoTotal As Date
      Dim strNaoConectado As String
      Dim strConectado As String
      Dim strTempoDeConexão As String
      Dim strTempoTotalConectado As String
      'Declare a user-defined variable to pass to the Shell_NotifyIcon
      'function.
      Private Type NOTIFYICONDATA
         cbSize As Long
         hWnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type
      'Declare the constants for the API function.
      'These constants can befound in the header file Shellapi.h.
      'The following constants are the messages sent to the
      'Shell_NotifyIcon function to add, modify, or delete an icon from the
      'taskbar status area.
      Private Const NIM_ADD = &H0
      Private Const NIM_MODIFY = &H1
      Private Const NIM_DELETE = &H2

      'The following constant is the message sent when a mouse event occurs
      'within the rectangular boundaries of the icon in the taskbar status
      'area.
      Private Const WM_MOUSEMOVE = &H200

      'The following constants are the flags that indicate the valid
      'members of the NOTIFYICONDATA data type.
      Private Const NIF_MESSAGE = &H1
      Private Const NIF_ICON = &H2
      Private Const NIF_TIP = &H4

      'The following constants are used to determine the mouse input on the
      'the icon in the taskbar status area.

      'Left-click constants.
      Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
      Private Const WM_LBUTTONDOWN = &H201     'Button down
      Private Const WM_LBUTTONUP = &H202       'Button up

      'Right-click constants.
      Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
      Private Const WM_RBUTTONDOWN = &H204     'Button down
      Private Const WM_RBUTTONUP = &H205       'Button up

      'Declare the API function call.
      Private Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

      'Dimension a variable as the user-defined data type.
      Dim nid As NOTIFYICONDATA
      ' Final das declarações necessárias para a função Colocar Icone na Bandeja"

'--------
Const TEMPO_DE_VERIFICAÇÃO_DA_CONEXÃO As Integer = 1000 ' 1.000 milisegundos = 1 segundo
Dim dtaDataHoraUltimaConexao As Date
Dim FirstTime As Boolean
Dim lngStatus As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'
Dim laststausOn As Boolean
Dim connect As Boolean
Private Sub cmdClearHistory_Click()
Dim i As Integer
List1.Clear
End Sub

Private Sub cmdHide_Click()
frmSessionController.Hide

End Sub

Private Sub cmdNewTime_Click()
frmNewTime.Show 1
End Sub


Private Sub Form_Load()
    frmSessionController.Hide
    Me.Caption = "Session Controller V " + Str(App.Major) + "." + Format(App.Minor, "0") + "." + Format(App.Revision, "00")
    Call ColocaIconeNaBandeja
    FirstTime = True
    strtabDescP(1) = "PGMAtivado ...."
    strtabDescP(2) = "Conectado ....."
    strtabDescP(3) = "Desconectado .."
    strtabDescP(4) = "Sessão rompida "
    strtabDescP(5) = "PGMInativado .."
    strTabDescE(1) = "PGMActived  ..."
    strTabDescE(2) = "Connected ....."
    strTabDescE(3) = "Disconnected .."
    strTabDescE(4) = "Interrupted ..."
    strTabDescE(5) = "PGMInactived .."
    intTimeOut = GetSetting(App.EXEName, "Configurações", "TimeOut", TEMPO_INICIAL_DE_TIMEOUT)
    strLanguage = GetSetting(App.EXEName, "Configurações", "Language", Português)
    Call AcertaALinguagem
    intTempoTimeOut = intTimeOut
    strStatus = "Inicial"
    Set dbdLog = OpenDatabase(App.Path + IIf(Right(App.Path, 1) = "\", "", "\") & "\NetLog.mdb")
    Call CarregaTabela
    Label2.BackColor = RGB(192, 224, 255)
    Label1.BackColor = RGB(192, 224, 255)
    Timer1.Interval = TEMPO_DE_VERIFICAÇÃO_DA_CONEXÃO
' Grava registro da data de ativação do programa ...
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          'Event occurs when the mouse pointer is within the rectangular
          'boundaries of the icon in the taskbar status area.
          Dim msg As Long
          Dim sFilter As String
          msg = X / Screen.TwipsPerPixelX
          Select Case msg
             Case WM_LBUTTONDOWN
             Case WM_LBUTTONUP
             Case WM_LBUTTONDBLCLK
                 Call MostraFormPrincipal
             Case WM_RBUTTONDOWN
                 PopupMenu PopMenu01
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          Case Else
          End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = 0 Then
    If MsgBox(strMsgFinal, vbDefaultButton2 + vbQuestion + vbYesNo, strMsgConfirmação) = vbYes Then
       Call EncerraPrograma
    Else
        Cancel = True
    End If
  Else
    Call EncerraPrograma
  End If
End Sub

Public Sub ColocaIconeNaBandeja()
         'Click this button to add an icon to the taskbar status area.

         'Set the individual values of the NOTIFYICONDATA data type.
         nid.cbSize = Len(nid)
         nid.hWnd = Me.hWnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = Me.Icon
         nid.szTip = Me.Caption & vbNullChar

         'Call the Shell_NotifyIcon function to add the icon to the taskbar
         'status area.
         Shell_NotifyIcon NIM_ADD, nid

End Sub

Private Sub mnuEncerrar_Click()
    If MsgBox(strMsgFinal, vbDefaultButton2 + vbQuestion + vbYesNo, strMsgConfirmação) = vbYes Then
       Call EncerraPrograma
    End If
End Sub

Private Sub mnuTelaInicial_Click()
   Call MostraFormPrincipal
End Sub

Private Sub Timer1_Timer()
Dim i As Integer, dtaTempoInicial As Date
Dim dtaNow As Date
    dtaNow = Now
    If strStatus = "Inicial" Then
       If tbLog.RecordCount <> 0 Then
         'Se existem fotos relacionadas ao álbum, carrega o gride, as primeiras 12 miniaturas e a primeira foto
         Do Until tbLog.EOF
            If tbLog!Id = 2 Then
               dtaTempoInicial = tbLog!DateTime
            ElseIf tbLog!Id = 3 Or tbLog!Id = 4 Or tbLog!Id = 5 Then
               If dtaTempoInicial <> 0 Then
                  dtaTempoTotal = dtaTempoTotal + (CDate(tbLog!DateTime) - dtaTempoInicial)
                  dtaTempoInicial = 0
               End If
            ElseIf tbLog!Id = 1 Then
               dtaTempoInicial = 0
            End If
            
            List1.AddItem tbLog!Desc & Format(tbLog!DateTime, "dd/mm/yyyy hh:mm:ss")
            List1.ListIndex = frmSessionController.List1.NewIndex
            tbLog.MoveNext
         Loop
       End If
       Call InsereEventoNoLog(1)
    End If
    If IsConnected Then
        Label2.Caption = strConectado
        If blnConnected Then
            lblTempoConexão.Caption = strTempoDeConexão & Format(dtaNow - dtaDataHoraUltimaConexao, "hh:mm:ss")
            lblTempoConexão.Visible = True
            lblTempoTotal.Caption = strTempoTotalConectado & Format(dtaTempoTotal + (dtaNow - dtaDataHoraUltimaConexao), "hh:mm:ss")
            lblTempoTotal.Visible = True
            lblTimeOut.Caption = " Tempo restante: " & intTimeOut & " minutos "
            lblTimeOut.Visible = True
            Me.Label1.BackColor = RGB(0, 0, 255)
            Me.Label2.BackColor = RGB(0, 0, 255)
            Me.Label2.ForeColor = RGB(0, 255, 0)
        Else
            intTimeOut = intTempoTimeOut
            Me.tmrConect.Enabled = True
        End If
        If strStatus <> "On" Then
            dtaDataHoraUltimaConexao = dtaNow
            Call InsereEventoNoLog(2)
        End If
        blnConnected = True
        strStatus = "On"
    Else
        lblTimeOut.Visible = False
        lblTempoConexão.Visible = False
        Label2.Caption = strNaoConectado
        blnConnected = False
        Me.tmrConect.Enabled = False
        If strStatus = "On" Then
           dtaTempoTotal = dtaTempoTotal + (dtaNow - dtaDataHoraUltimaConexao)
           Call InsereEventoNoLog(3)
           Label2.ForeColor = RGB(255, 0, 0)
           Label1.BackColor = RGB(192, 224, 255)
           Label2.BackColor = RGB(192, 224, 255)
        End If
        strStatus = "Off"
        lblTempoTotal.Caption = strTempoTotalConectado & Format(dtaTempoTotal, "hh:mm:ss")
        lblTempoTotal.Visible = True
    End If

End Sub

Private Sub tlbFunções_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
   Case "NewTimeOut"
     'Set new TimeOut
     frmNewTime.Show 1
   Case "ResumoMensal"
     'Lista o resumo das conexões dia a dia do último mês ...
     frmResumo.Show 1
   Case "Esconder"
     'Esconde o Form principal e fica só na bandeja ...
     Me.Hide
   Case "Lingua"
     'Altera a lingua utilizada de português p/ inglês e vice-versa ...
   If MsgBox(strMsgLinguagem, vbDefaultButton2 + vbQuestion + vbYesNo, strMsgConfirmação) = vbYes Then
     strLanguage = IIf(strLanguage = Português, English, Português)
     SaveSetting App.EXEName, "Configurações", "Language", strLanguage
     Call AcertaALinguagem
   End If
   Case "Sair"
    If MsgBox(strMsgFinal, vbDefaultButton2 + vbQuestion + vbYesNo, strMsgConfirmação) = vbYes Then
        Call EncerraPrograma
     End If
End Select
End Sub

Private Sub tmrConect_Timer()
Dim i As Integer
'Debug.Print Now & " - tmrConect: " & intTimeOut
If blnConnected Then
   intTimeOut = intTimeOut - 1
   If intTimeOut = 0 Then
       blnEncerraConexão = False
       frmTimeOut.Show
       frmSessionController.tmrConect.Enabled = False
   End If
End If
End Sub

Public Sub CarregaTabela()
       Dim datHoje As Date, blnTabelaOK As Boolean
       datHoje = Now
       Set tbLog = dbdLog.OpenRecordset("SELECT * FROM Log WHERE Dia='" & Format(datHoje, "dd") & "'")
       'tbLog.MoveFirst
       If tbLog.EOF <> True And tbLog.BOF <> True Then
          If Format(tbLog!DateTime, "yyyymm") <> Format(datHoje, "yyyymm") Then
             dbdLog.Execute "DELETE * FROM Log WHERE Dia='" & Format(datHoje, "dd") & "'"
             Set tbLog = dbdLog.OpenRecordset("SELECT * FROM Log WHERE Dia='" & Format(datHoje, "dd") & "'")
          End If
       End If
End Sub

Public Sub MostraFormPrincipal()
                Me.Caption = Me.Caption
                Me.Show
                AlteraPosiçãoDoForm Me, "Topo"
                AlteraPosiçãoDoForm Me, "Normal"
                List1.ListIndex = List1.NewIndex
End Sub

Public Sub EncerraPrograma()
        Call InsereEventoNoLog(5)
        tbLog.Close
        'Delete the added icon from the taskbar status area when the
        'program ends.
        Shell_NotifyIcon NIM_DELETE, nid
        End
End Sub

Private Sub AcertaALinguagem()
Dim intI As Integer
Select Case strLanguage
   Case Português
      For intI = 1 To 5
         strTabDesc(intI) = strtabDescP(intI)
      Next intI
      Label1.Caption = "Situação atual:"
      Me.mnuEncerrar.Caption = "Ence&rrar"
      Me.mnuTelaInicial.Caption = "&Mostrar painel principal ..."
      Me.tlbFunções.Buttons(1).ToolTipText = "Indicar novo timeout para confirmação ..."
      Me.tlbFunções.Buttons(2).ToolTipText = "Lista conexões dos últimos 30 dias ..."
      Me.tlbFunções.Buttons(3).ToolTipText = "Esconde o painel principal ..."
      Me.tlbFunções.Buttons(4).ToolTipText = "Português --> Inglês .../Portuguese --> English ..."
      Me.tlbFunções.Buttons(5).ToolTipText = "Encerra o programa ..."
      strMsgErro01 = "Erro na verificação da conexão Internet - CC="
      strTempoDeConexão = "Tempo de conexão: "
      strTempoTotalConectado = " Tempo total conectado: "
      strConectado = "Conectado"
      strNaoConectado = "Não Conectado"
      strMsgFinal = "Tem certeza que deseja encerrar agora ?"
      strMsgConfirmação = "Confirmação:"
      strMsgLinguagem = "Tem certeza que quer mudar o idioma para o Inglês ?" & Chr(10) & Chr(13) & "(Are you sure you want to exchange the language to English ?)"
   Case English
      For intI = 1 To 5
         strTabDesc(intI) = strTabDescE(intI)
      Next intI
      Label1.Caption = "Status now:"
      Me.mnuEncerrar.Caption = "E&xit"
      Me.mnuTelaInicial.Caption = "&Show main panel ..."
      Me.tlbFunções.Buttons(1).ToolTipText = "New timeout setting ..."
      Me.tlbFunções.Buttons(2).ToolTipText = "View last 30 days connections ..."
      Me.tlbFunções.Buttons(3).ToolTipText = "Hide the main panel ..."
      Me.tlbFunções.Buttons(4).ToolTipText = "English --> Portuguese .../Inglês --> Português..."
      Me.tlbFunções.Buttons(5).ToolTipText = "Quit the program ..."
      strMsgErro01 = "Error on Internet connection verification - CC="
      strTempoDeConexão = "Connection Time: "
      strTempoTotalConectado = " Connected total time: "
      strConectado = "Connected"
      strNaoConectado = "Not Connected"
      strMsgFinal = "Are you sure you want to finish now ?"
      strMsgConfirmação = "Confirmation:"
      strMsgLinguagem = "Are you sure you want to exchange the language to Portuguese ?" & Chr(10) & Chr(13) & "(Tem certeza que quer mudar o idioma para o Português ?)"
End Select
End Sub
