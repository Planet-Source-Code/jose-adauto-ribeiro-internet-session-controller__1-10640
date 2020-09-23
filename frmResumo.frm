VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmResumo 
   Caption         =   "x"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxGrid 
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8070
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Menu mnuRetornar 
      Caption         =   "&Retornar"
   End
End
Attribute VB_Name = "frmResumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim tbLogMensal As Recordset
Dim strTempoTotal As String
Public strDiaSelecionado As String


Private Sub flxGrid_DblClick()
strDiaSelecionado = flxGrid.TextMatrix(flxGrid.Row, 0)
frmDiario.Show 1

End Sub

Private Sub Form_Load()
Dim strMesAtual As String, strMesAnterior As String
Dim intMesAnterior As Integer, intAnoAnterior As Integer
Dim intLinha As Integer, strDiaAnterior As String
Dim dtaTempoInicial As Date, dtaTempoTotal As Date
Dim blnTemAlgo As Boolean, dtaTotalDoPeriodo As Date
   Call AcertaALinguagem
   If Format(Now, "mm") = "01" Then
      intMesAnterior = 11
      intAnoAnterior = CInt(Format(Now, "yyyy")) - 1
   Else
      intAnoAnterior = CInt(Format(Now, "yyyy"))
      intMesAnterior = CInt(Format(Now, "mm")) - 1
   End If
   strMesAtual = Format(Now, "/mm/yyyy")
   strMesAnterior = "/" & Format(intMesAnterior, "00") & "/" & Format(intAnoAnterior, "0000")
   Set tbLogMensal = dbdLog.OpenRecordset("SELECT * FROM Log ORDER BY DateTime DESC")
   'Set tbLogMensal = dbdLog.OpenRecordset("SELECT * FROM Log WHERE DateTime='%%" & strMesAtual & "' OR DateTime='" & strMesAnterior & "'")
   flxGrid.Cols = 2
   flxGrid.ColWidth(0) = 1200
   flxGrid.ColWidth(1) = 2000
   flxGrid.ColAlignment(0) = flexAlignCenterCenter
   flxGrid.ColAlignment(1) = flexAlignCenterCenter
   flxGrid.TextMatrix(0, 0) = IIf(strLanguage = Português, "Data", "Date")
   flxGrid.TextMatrix(0, 1) = IIf(strLanguage = Português, "Tempo de Conexão", "Connection Time")
   flxGrid.Rows = 1
   strDiaAnterior = ""
   If Not tbLogMensal.EOF Then
      strDiaAnterior = tbLogMensal!dia
   End If
   Do Until tbLogMensal.EOF
      If strDiaAnterior = tbLogMensal!dia Then
            blnTemAlgo = True
            strMesAnterior = Format(tbLogMensal!DateTime, "dd/mm/yyyy")
            If tbLogMensal!Id = 2 Then
               dtaTempoInicial = tbLogMensal!DateTime
            ElseIf tbLogMensal!Id = 3 Or tbLogMensal!Id = 4 Or tbLogMensal!Id = 5 Then
               If dtaTempoInicial <> 0 Then
                  dtaTempoTotal = dtaTempoTotal + (CDate(tbLogMensal!DateTime) - dtaTempoInicial)
                  dtaTempoInicial = 0
               End If
            ElseIf tbLogMensal!Id = 1 Then
               dtaTempoInicial = 0
            End If
            tbLogMensal.MoveNext
      Else
            'blnTemAlgo = False
            flxGrid.AddItem strMesAnterior & vbTab & Format(dtaTempoTotal, "hh:mm:ss")
            dtaTotalDoPeriodo = dtaTotalDoPeriodo + dtaTempoTotal
            dtaTempoInicial = 0
            dtaTempoTotal = 0
            strDiaAnterior = tbLogMensal!dia
            tbLogMensal.MoveNext
      End If
   Loop
   If blnTemAlgo Then
      flxGrid.AddItem strMesAnterior & vbTab & Format(dtaTempoTotal, "hh:mm:ss")
      lblTotal.Caption = strTempoTotal & Format(dtaTotalDoPeriodo, "hh:mm:ss")
      lblTotal.Visible = True
   End If
End Sub

Private Sub mnuRetornar_Click()
Unload Me
End Sub
Private Sub AcertaALinguagem()
Select Case strLanguage
   Case Português
      Me.Caption = "Resumo dos últimos 30 dias"
      strTempoTotal = " Tempo total: "
      Me.mnuRetornar.Caption = "&Retornar"
      Me.flxGrid.ToolTipText = "Para ver detalhes do dia, dê duplo clique na linha correspondente."
   Case English
      Me.Caption = "Last 30 days summary"
      strTempoTotal = " Total time: "
      Me.mnuRetornar.Caption = "&Return"
      Me.flxGrid.ToolTipText = "To view details of the day, double click on the respective row."
End Select
End Sub

