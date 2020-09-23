Attribute VB_Name = "basSessionController"

Public Const TEMPO_INICIAL_DE_TIMEOUT As Integer = 30  ' 30 minutos

Public dbdLog As Database
Public tbLog As Recordset
Public strTabDesc(1 To 5) As String
Public strTabDescE(1 To 5) As String
Public strtabDescP(1 To 5) As String
Public strLanguage As String
Global Const Português As String = "P", English As String = "E"
Public intTempoTimeOut As Integer
Public strMsgErro01 As String

Public Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lpRasCon As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasCon As Long, lpStatus As Any) As Long
Public Declare Function RasHangUp Lib "RasApi32.dll" Alias "RasHangUpA" (ByVal hRasCon As Long) As Long


Public blnConnected As Boolean
Public AlterouTimeout As Boolean
Public intTimeOut As Integer
Public blnEncerraConexão As Boolean
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags&, ByVal wReserved&)
Global Const EWX_FORCE = 4 'constants needed for exiting Windows
Global Const EWX_LOGOFF = 0
Global Const EWX_REBOOT = 2
Global Const EWX_SHUTDOWN = 1
'* Variaveis utilizadas para a função que coloca o
'* form no Topo ou retira essa opção
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2


Public Const RAS95_MaxEntryName = 256
Public Const RAS95_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 32
'
Public RetCode As Long


Public Type RASCONN95
    dwSize As Long
    hRasCon As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
'
Public Type RASCONNSTATUS95
    dwSize As Long
    RasConnState As Long
    dwError As Long
    szDeviceType(RAS95_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type

Public TRasCon(255) As RASCONN95

Public Function AlteraPosiçãoDoForm(frm As Form, Posição As String)
    Select Case Posição
       Case "Topo"
          'To set Form1 as a TopMost form, do the following:
          AlteraPosiçãoDoForm = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
       Case "Normal"
         'To turn off topmost (make the form act normal again):
         AlteraPosiçãoDoForm = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End Select
End Function
Public Function IsConnected() As Boolean
'Dim TRasCon(255) As RASCONN95
Dim lg As Long
Dim lpcon As Long
Dim RetVal As Long
Dim Tstatus As RASCONNSTATUS95
'
TRasCon(0).dwSize = 412
lg = 256 * TRasCon(0).dwSize

RetVal = RasEnumConnections(TRasCon(0), lg, lpcon)
If RetVal <> 0 Then
   MsgBox "SessionController.900I " & strMsgErro01 & RetVal
   Exit Function
End If
'
Tstatus.dwSize = 160
RetVal = RasGetConnectStatus(TRasCon(0).hRasCon, Tstatus)
If blnEncerraConexão And Not IsMissing(RetCode) Then
    RetCode = Tstatus.RasConnState
End If
If Tstatus.RasConnState = &H2000 Then
   IsConnected = True
Else
   IsConnected = False
End If
 
End Function


Public Sub InsereEventoNoLog(Id As Long)
        Dim datDate As Date
        datDate = Now
        tbLog.AddNew
        tbLog!Id = Id
        tbLog!Desc = strTabDesc(Id)
        tbLog!DateTime = datDate
        tbLog!dia = Format(datDate, "dd")
        tbLog.Update
        frmSessionController.List1.AddItem strTabDesc(Id) & Format(datDate, "dd/mm/yyyy hh:mm:ss")
        frmSessionController.List1.ListIndex = frmSessionController.List1.NewIndex

End Sub
