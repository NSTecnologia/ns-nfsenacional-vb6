Attribute VB_Name = "NSNFSeNacional"
'==============================================================================
' NSNFSeNacional - Comunicação com a NS API para NFSe Nacional (VB6).
'
' Equivalente ao exemplo ns-nfsenacional-vbnet:
'   emitirNFSeSincrono -> envia, consulta o status e baixa o documento.
'
' Dependências: somente componentes do Windows
'   - WinHttp.WinHttpRequest.5.1 (winhttp.dll)
'   - Scripting.Dictionary      (scrrun.dll)
'==============================================================================
Option Explicit

' Defina aqui o token de acesso fornecido pela NS Tecnologia
Private Const TOKEN As String = "Seu_Token_Aqui"

' Leiaute da DPS enviado no header "nt" da emissão:
'   "nt009" -> leiaute novo (NT 009 / RTC IBS-CBS), exigido pela Receita.
'              finNFSe vai na raiz do infDPS, antes de cLocEmi.
'   ""      -> não envia o header (leiaute antigo, NT 004).
Public Const VERSAO_NT As String = "nt009"

' Endpoints NS API
Public Const URL_ENVIO As String = "https://nfsenacional.ns.eti.br/nfse/issue"
Public Const URL_CONS_STATUS As String = "https://nfsenacional.ns.eti.br/nfse/issue/status"
Public Const URL_DOWNLOAD As String = "https://nfsenacional.ns.eti.br/nfse/get"

' Tempo de espera entre o envio e a consulta (ms)
Public Const TEMPO_ESPERA As Long = 400

' WinHttp: TLS 1.0 (&H80) + TLS 1.1 (&H200) + TLS 1.2 (&H800)
Private Const WINHTTP_OPTION_SECURE_PROTOCOLS As Long = 9
Private Const WINHTTP_FLAG_TLS As Long = &HA80

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'------------------------------------------------------------------------------
' Envia um conteúdo via POST para a API e devolve o texto da resposta.
' Em falha de comunicação devolve "" e grava o erro no log.
'------------------------------------------------------------------------------
Public Function enviaConteudoParaAPI(ByVal conteudo As String, ByVal url As String, ByVal tpConteudo As String, _
                                     Optional ByVal nt As String = "") As String
    Dim http As Object
    Dim resposta() As Byte
    Dim corpo As Variant
    Dim statusHttp As Long

    On Error GoTo Falha

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False

    ' Força TLS 1.2 em Windows mais antigos (ignora se não suportado)
    On Error Resume Next
    http.Option(WINHTTP_OPTION_SECURE_PROTOCOLS) = WINHTTP_FLAG_TLS
    On Error GoTo Falha

    http.SetTimeouts 30000, 30000, 60000, 60000
    http.SetRequestHeader "X-AUTH-TOKEN", TOKEN
    If Len(nt) > 0 Then http.SetRequestHeader "nt", nt

    Select Case LCase$(tpConteudo)
        Case "txt"
            http.SetRequestHeader "Content-Type", "text/plain;charset=utf-8"
        Case "xml"
            http.SetRequestHeader "Content-Type", "application/xml;charset=utf-8"
        Case Else
            http.SetRequestHeader "Content-Type", "application/json;charset=utf-8"
    End Select

    ' Envia o corpo em UTF-8
    corpo = textoParaUtf8(conteudo)
    http.Send corpo

    statusHttp = http.Status
    If IsArray(http.ResponseBody) Then
        resposta = http.ResponseBody
        enviaConteudoParaAPI = utf8ParaTexto(resposta)
    End If

    Select Case statusHttp
        Case 401
            MsgBox "Token não enviado ou inválido", vbExclamation
        Case 403
            MsgBox "Token sem permissão", vbExclamation
        Case 404
            MsgBox "Não encontrado, verifique o retorno para mais informações", vbExclamation
    End Select

    Exit Function

Falha:
    gravarLinhaLog "NFSE: ", "[ERRO_COMUNICACAO]"
    gravarLinhaLog "NFSE: ", Err.Number & " - " & Err.Description
    enviaConteudoParaAPI = ""
End Function

'------------------------------------------------------------------------------
' Envio da DPS (json, xml ou txt)
'------------------------------------------------------------------------------
Public Function emitirDocumento(ByVal conteudo As String, ByVal tpConteudo As String, _
                                ByVal cnpjEmitente As String, Optional ByVal a3 As Boolean = False) As String
    Dim resposta As String

    gravarLinhaLog "NFSE: ", "[ENVIO_DADOS]"
    gravarLinhaLog "NFSE: ", conteudo

    resposta = enviaConteudoParaAPI(conteudo, URL_ENVIO, tpConteudo, VERSAO_NT)

    gravarLinhaLog "NFSE: ", "[ENVIO_RESPOSTA]"
    gravarLinhaLog "NFSE: ", resposta

    emitirDocumento = resposta
End Function

'------------------------------------------------------------------------------
' Consulta do status de processamento pelo nsNRec
'------------------------------------------------------------------------------
Public Function consultarStatusProcessamento(ByVal CNPJ As String, ByVal nsNRec As String, ByVal tpAmb As String) As String
    Dim json As String
    Dim resposta As String

    json = "{" & JsonPar("CNPJ", CNPJ) & "," & _
                 JsonPar("nsNRec", nsNRec) & "," & _
                 JsonPar("tpAmb", tpAmb) & "}"

    gravarLinhaLog "NFSE: ", "[CONSULTA_DADOS]"
    gravarLinhaLog "NFSE: ", json

    resposta = enviaConteudoParaAPI(json, URL_CONS_STATUS, "json")

    gravarLinhaLog "NFSE: ", "[CONSULTA_RESPOSTA]"
    gravarLinhaLog "NFSE: ", resposta

    consultarStatusProcessamento = resposta
End Function

'------------------------------------------------------------------------------
' Download do documento (tpDown: X = XML, P = PDF, XP = ambos)
'------------------------------------------------------------------------------
Public Function downloadDocumento(ByVal chDPS As String, ByVal chNFSe As String, ByVal CNPJ As String, _
                                  ByVal tpAmb As String, ByVal tpDown As String) As String
    Dim json As String
    Dim resposta As String
    Dim status As String

    json = "{" & JsonPar("chDPS", chDPS) & "," & _
                 JsonPar("chNFSe", chNFSe) & "," & _
                 JsonPar("tpAmb", tpAmb) & "," & _
                 JsonPar("CNPJ", CNPJ) & "," & _
                 JsonPar("tpDown", tpDown) & "}"

    gravarLinhaLog "NFSE: ", "[DOWNLOAD_DADOS]"
    gravarLinhaLog "NFSE: ", json

    resposta = enviaConteudoParaAPI(json, URL_DOWNLOAD, "json")
    status = JsonStr(JsonParse(resposta), "status")

    ' Em caso de sucesso grava só o status, para não encher o log com o PDF
    If status <> "200" And status <> "100" Then
        gravarLinhaLog "NFSE: ", "[DOWNLOAD_RESPOSTA]"
        gravarLinhaLog "NFSE: ", resposta
    Else
        gravarLinhaLog "NFSE: ", "[DOWNLOAD_STATUS]"
        gravarLinhaLog "NFSE: ", status
    End If

    downloadDocumento = resposta
End Function

'------------------------------------------------------------------------------
' Download do documento e gravação dos arquivos em disco
'------------------------------------------------------------------------------
Public Function downloadDocumentoESalvar(ByVal chDPS As String, ByVal chNFSe As String, ByVal CNPJ As String, _
                                         ByVal tpAmb As String, ByVal tpDown As String, _
                                         ByVal caminho As String, ByVal nome As String, _
                                         Optional ByVal exibeNaTela As Boolean = False) As String
    Dim resposta As String
    Dim ret As Object
    Dim status As String

    resposta = downloadDocumento(chDPS, chNFSe, CNPJ, tpAmb, tpDown)
    downloadDocumentoESalvar = resposta

    Set ret = JsonParse(resposta)
    status = JsonStr(ret, "status")
    If status <> "200" And status <> "100" Then Exit Function

    If Not criarPasta(caminho) Then
        gravarLinhaLog "NFSE: ", "[CRIAR_DIRETORIO]" & caminho
        MsgBox "Não foi possível criar a pasta: " & caminho, vbCritical
        Exit Function
    End If
    If Right$(caminho, 1) <> "\" Then caminho = caminho & "\"

    If InStr(UCase$(tpDown), "X") > 0 Then
        salvarXML JsonStr(ret, "xml"), caminho, nome
    End If

    If InStr(UCase$(tpDown), "P") > 0 Then
        salvarPDF JsonStr(ret, "pdfDocumento"), caminho, nome
        If exibeNaTela Then abrirArquivo caminho & nome & ".pdf"
    ElseIf InStr(UCase$(tpDown), "X") = 0 Then
        MsgBox "Ocorreu um erro, veja o retorno da API para mais informações", vbExclamation
    End If
End Function

'------------------------------------------------------------------------------
' Emissão síncrona: envio + consulta de status + download.
' Retorna um JSON com: statusEnvio, statusConsulta, statusDownload, cStat,
' chNFSe, chDPS, nProt, motivo, nsNRec, erros.
'------------------------------------------------------------------------------
Public Function emitirNFSeSincrono(ByVal conteudo As String, ByVal tpConteudo As String, ByVal CNPJ As String, _
                                   ByVal tpDown As String, ByVal tpAmb As String, ByVal caminho As String, _
                                   Optional ByVal exibeNaTela As Boolean = False, _
                                   Optional ByVal a3 As Boolean = False) As String
    Dim statusEnvio As String
    Dim statusConsulta As String
    Dim statusDownload As String
    Dim motivo As String
    Dim nsNRec As String
    Dim chDPS As String
    Dim chNFSe As String
    Dim cStat As String
    Dim nProt As String
    Dim resposta As String
    Dim retorno As String
    Dim erros As Object
    Dim ret As Object
    Dim i As Long

    gravarLinhaLog "NFSE: ", "[EMISSAO_SINCRONA_INICIO]"

    '--- Envio -----------------------------------------------------------------
    resposta = emitirDocumento(conteudo, tpConteudo, CNPJ, a3)
    Set ret = JsonParse(resposta)
    statusEnvio = JsonStr(ret, "status")

    If statusEnvio = "200" Or statusEnvio = "-6" Then
        nsNRec = JsonStr(ret, "nsNRec")

        Sleep TEMPO_ESPERA

        '--- Consulta (repete enquanto a SEFAZ responder 997 - em processamento)
        For i = 0 To 2
            resposta = consultarStatusProcessamento(CNPJ, nsNRec, tpAmb)
            Set ret = JsonParse(resposta)

            If JsonStr(ret, "status") = "-2" And JsonStr(ret, "erro.cStat") = "997" Then
                Sleep TEMPO_ESPERA
            Else
                Exit For
            End If
        Next

        statusConsulta = JsonStr(ret, "status")

        If statusConsulta = "200" Then
            cStat = JsonStr(ret, "cStat")
            motivo = JsonStr(ret, "xMotivo")

            If cStat = "100" Or cStat = "150" Then
                chDPS = JsonStr(ret, "chDPS")
                chNFSe = JsonStr(ret, "chNFSe")
                nProt = JsonStr(ret, "nProt")

                '--- Download --------------------------------------------------
                resposta = downloadDocumentoESalvar(chDPS, chNFSe, CNPJ, tpAmb, tpDown, _
                                                    caminho, chNFSe & "-procNFSe", exibeNaTela)
                Set ret = JsonParse(resposta)
                statusDownload = JsonStr(ret, "status")
                If statusDownload <> "200" Then motivo = JsonStr(ret, "motivo")
            End If
        Else
            motivo = JsonStr(ret, "motivo")
            Set erros = JsonObj(ret, "erros")
        End If

    ElseIf statusEnvio = "-7" Then
        motivo = JsonStr(ret, "motivo")
        nsNRec = JsonStr(ret, "nsNRec")

    ElseIf statusEnvio = "-4" Then
        motivo = JsonStr(ret, "motivo")
        Set erros = JsonObj(ret, "erros")

    ElseIf statusEnvio = "-9" Then
        motivo = JsonStr(ret, "erro.xMotivo")
        cStat = JsonStr(ret, "erro.cStat")

    Else
        motivo = JsonStr(ret, "motivo")
        If Len(motivo) = 0 Then motivo = resposta
        If Len(motivo) = 0 Then motivo = "Sem resposta da API, verifique o log"
    End If

    '--- Monta o JSON de retorno ----------------------------------------------
    retorno = "{" & _
        JsonPar("statusEnvio", statusEnvio) & "," & _
        JsonPar("statusConsulta", statusConsulta) & "," & _
        JsonPar("statusDownload", statusDownload) & "," & _
        JsonPar("cStat", cStat) & "," & _
        JsonPar("chNFSe", chNFSe) & "," & _
        JsonPar("chDPS", chDPS) & "," & _
        JsonPar("nProt", nProt) & "," & _
        JsonPar("motivo", motivo) & "," & _
        JsonPar("nsNRec", nsNRec) & "," & _
        """erros"":" & JsonStringify(erros) & "}"

    gravarLinhaLog "NFSE: ", "[JSON_RETORNO]"
    gravarLinhaLog "NFSE: ", retorno
    gravarLinhaLog "NFSE: ", "[EMISSAO_SINCRONA_FIM]"

    emitirNFSeSincrono = retorno
End Function
