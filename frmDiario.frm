VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDiario 
   Caption         =   "Resumo diário "
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   3990
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flxGrid 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   7646
      _Version        =   393216
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDia 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Detalhamento do dia: dd/aa/aaaa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mnuRetornar 
      Caption         =   "&Retornar"
   End
End
Attribute VB_Name = "frmDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMsgDetalhe As String
Private Sub Form_Load()
Dim tbLogDiario As Recordset
   frmResumo.Hide
   Me.mnuRetornar.Caption = IIf(strLanguage = Português, "&Retornar", "&Return")
   Me.Caption = IIf(strLanguage = Português, "Resumo Diário", "Diary Summary")
   lblDia.Caption = IIf(strLanguage = Português, "Detalhamento do dia: ", "Details of day: ") & frmResumo.strDiaSelecionado
   Set tbLogDiario = dbdLog.OpenRecordset("SELECT * FROM Log WHERE Dia = '" & Format(frmResumo.strDiaSelecionado, "dd") & "' ORDER BY DateTime")
   flxGrid.Cols = 2
   flxGrid.ColWidth(0) = 1800
   flxGrid.ColWidth(1) = 1200
   flxGrid.ColAlignment(0) = flexAlignCenterCenter
   flxGrid.ColAlignment(1) = flexAlignCenterCenter
   flxGrid.TextMatrix(0, 1) = IIf(strLanguage = Português, "Hora", "Hour")
   flxGrid.TextMatrix(0, 0) = IIf(strLanguage = Português, "Evento", "Event")
   flxGrid.Rows = 1

   Do Until tbLogDiario.EOF

      flxGrid.AddItem tbLogDiario!Desc & vbTab & Format(tbLogDiario!DateTime, "hh:mm:ss")
      tbLogDiario.MoveNext
   Loop
End Sub

Private Sub mnuRetornar_Click()
Unload Me
frmResumo.Show 1
End Sub
