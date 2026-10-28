VERSION 5.00
Begin VB.Form frmPrincipal
   Caption         =   "NS API - NFSe Nacional (exemplo VB6)"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRetorno
      BeginProperty Font
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Top             =   3360
      Width           =   9375
   End
   Begin VB.CommandButton cmdEmitir
      Caption         =   "Emitir NFSe (síncrono)"
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CheckBox chkExibe
      Caption         =   "Abrir o PDF após a autorização"
      Height          =   315
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.TextBox txtCaminho
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   7575
   End
   Begin VB.TextBox txtTpDown
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Text            =   "XP"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtTpAmb
      Height          =   315
      Left            =   1920
      TabIndex        =   5
      Text            =   "2"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtCNPJ
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   2655
   End
   Begin VB.TextBox txtArquivo
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label lblRetorno
      Caption         =   "Retorno:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3060
      Width           =   1695
   End
   Begin VB.Label lblCaminho
      Caption         =   "Pasta das notas:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2070
      Width           =   1695
   End
   Begin VB.Label lblTpDown
      Caption         =   "tpDown (X/P/XP):"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   1695
   End
   Begin VB.Label lblTpAmb
      Caption         =   "tpAmb (1/2):"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1110
      Width           =   1695
   End
   Begin VB.Label lblCNPJ
      Caption         =   "CNPJ emitente:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   1695
   End
   Begin VB.Label lblArquivo
      Caption         =   "Arquivo da DPS:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1695
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim pasta As String

    pasta = pastaApp()

    txtArquivo.Text = pasta & "nfse.json"
    txtCaminho.Text = pasta & "notas\"
End Sub

Private Sub cmdEmitir_Click()
    Dim conteudo As String
    Dim tpConteudo As String
    Dim retorno As String

    On Error GoTo Falha

    If Len(Dir$(txtArquivo.Text)) = 0 Then
        MsgBox "Arquivo não encontrado: " & txtArquivo.Text, vbExclamation
        Exit Sub
    End If

    ' O tipo do conteúdo segue a extensão do arquivo (json, xml ou txt)
    tpConteudo = LCase$(Mid$(txtArquivo.Text, InStrRev(txtArquivo.Text, ".") + 1))
    If tpConteudo <> "xml" And tpConteudo <> "txt" Then tpConteudo = "json"

    conteudo = lerArquivoUtf8(txtArquivo.Text)

    Screen.MousePointer = vbHourglass
    cmdEmitir.Enabled = False
    txtRetorno.Text = "Emitindo..."
    DoEvents

    retorno = NSNFSeNacional.emitirNFSeSincrono(conteudo, tpConteudo, txtCNPJ.Text, txtTpDown.Text, _
                                                txtTpAmb.Text, txtCaminho.Text, (chkExibe.Value = vbChecked))

    Screen.MousePointer = vbDefault
    cmdEmitir.Enabled = True
    txtRetorno.Text = retorno

    TratarRetorno retorno
    Exit Sub

Falha:
    Screen.MousePointer = vbDefault
    cmdEmitir.Enabled = True
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' Exemplo de tratamento do JSON devolvido por emitirNFSeSincrono
Private Sub TratarRetorno(ByVal retorno As String)
    Dim ret As Object
    Dim statusEnvio As String
    Dim statusConsulta As String
    Dim statusDownload As String
    Dim cStat As String
    Dim motivo As String
    Dim erros As String

    Set ret = JsonParse(retorno)
    statusEnvio = JsonStr(ret, "statusEnvio")
    statusConsulta = JsonStr(ret, "statusConsulta")
    statusDownload = JsonStr(ret, "statusDownload")
    cStat = JsonStr(ret, "cStat")
    motivo = JsonStr(ret, "motivo")
    If Not JsonObj(ret, "erros") Is Nothing Then erros = JsonStringify(JsonObj(ret, "erros"))

    ' Testa se houve sucesso no envio
    If statusEnvio = "200" Or statusEnvio = "-6" Then

        ' Testa se houve sucesso na consulta
        If statusConsulta = "200" Then

            ' Testa se a nota foi autorizada
            If cStat = "100" Or cStat = "150" Then
                MsgBox motivo, vbInformation

                ' Testa se o download teve problemas
                If statusDownload <> "200" Then
                    MsgBox "Erro no Download", vbExclamation
                End If
            Else
                MsgBox motivo, vbExclamation
            End If
        Else
            MsgBox motivo & vbCrLf & erros, vbExclamation
        End If
    Else
        MsgBox motivo & vbCrLf & erros, vbExclamation
    End If
End Sub
