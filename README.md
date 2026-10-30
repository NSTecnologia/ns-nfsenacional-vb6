# ns-nfsenacional-vb6

Utilizando a NS API, este exemplo - criado em Visual Basic 6 - possui funcionalidades para consumir o documento fiscal eletrônico:
+ NFSe Nacional;

É a versão VB6 do exemplo [ns-nfsenacional-vbnet](https://github.com/NSTecnologia/ns-nfsenacional-vbnet), com as mesmas funções e o mesmo JSON de retorno.

## Primeiros passos:

### Integrando ao sistema:

1. Baixe o projeto ns-nfsenacional-vb6 para sua máquina;
2. Adicione ao seu projeto VB6 (menu *Project > Add Module*) os três módulos:
   + **NSNFSeNacional.bas** - funções de comunicação com a API;
   + **Genericos.bas** - log, gravação de arquivos, UTF-8 e Base64;
   + **modJSON.bas** - leitura e geração de JSON.
3. Defina o seu **token** na constante `TOKEN` do módulo **NSNFSeNacional.bas**.

> **Leiaute NT 009 (IBS/CBS):** a emissão envia o header `nt: nt009` (constante `VERSAO_NT`), que faz a API gerar a DPS no leiaute novo exigido pela Receita. Nesse leiaute o campo `finNFSe` é obrigatório e vai na raiz do `infDPS`, logo após `tpEmit`. Sem o header, a API usa o leiaute antigo e a Receita rejeita com `E1235` (*"invalid child element 'cLocEmi' ... expected 'cMotivoEmisTI, finNFSe'"*).
>
> O conteúdo pode ser enviado em **JSON** ou em **XML** (`tpConteudo` = `json` ou `xml`); no formulário de exemplo o tipo segue a extensão do arquivo.

**Não é preciso instalar nenhuma DLL ou OCX.** O exemplo usa apenas componentes que já vêm no Windows:

Componente | Uso
:----------|:----
`WinHttp.WinHttpRequest.5.1` | Requisições HTTPS para a API
`Scripting.Dictionary` | Objetos JSON
`crypt32.dll` | Decodificação do PDF (Base64)
`kernel32.dll` | Conversão UTF-8

> **TLS 1.2:** no Windows 8.1, 10 e 11 funciona direto. No Windows 7 é necessário instalar a atualização [KB3140245](https://support.microsoft.com/kb/3140245) e habilitar o TLS 1.2 no WinHTTP.

O projeto **NSNFSeNacionalVB6.vbp** traz um formulário de exemplo (**frmPrincipal**) pronto para testar a emissão com o arquivo **nfse.json**.

------

## Emissão Síncrona:

### Realizando uma Emissão Síncrona:

Você poderá utilizar a função **emitirNFSeSincrono** do módulo NSNFSeNacional.

##### Parâmetros:

Parâmetros     | Descrição
:-------------:|:-----------
**conteudo**   | Conteúdo de emissão do documento.
**tpConteudo** | Tipo de conteúdo que está sendo enviado. Valores possíveis: json, xml, txt
**CNPJ**       | CNPJ do emitente do documento.
**tpDown**     | Tipo de arquivos a serem baixados. Valores possíveis: <ul> <li>**X** - XML</li> <li>**P** - PDF</li> <li>**XP** - XML e PDF</li> </ul>
**tpAmb**      | Ambiente onde foi autorizado o documento. Valores possíveis: <ul> <li>1 - produção</li> <li>2 - homologação</li> </ul>
**caminho**    | Caminho onde devem ser salvos os documentos baixados.
**exibeNaTela**| Se for baixado o PDF, exibe na tela após a autorização. Valores possíveis: <ul> <li>**True** - será exibido</li> <li>**False** - não será exibido (padrão)</li> </ul>

##### Exemplo de chamada:

    Dim conteudo As String
    Dim retorno As String

    conteudo = lerArquivoUtf8("C:\exemplos\nfse.json")
    retorno = NSNFSeNacional.emitirNFSeSincrono(conteudo, "json", "00000000000000", "XP", "2", "C:\exemplos\notas\", True)
    MsgBox retorno

A função **emitirNFSeSincrono** faz o envio, a consulta e o download do documento, utilizando as funções **emitirDocumento**, **consultarStatusProcessamento** e **downloadDocumentoESalvar**. Os arquivos são gravados como `<chNFSe>-procNFSe.xml` e `<chNFSe>-procNFSe.pdf`.

##### Exemplo de tratamento de retorno:

O JSON retornado terá os campos: statusEnvio, statusConsulta, statusDownload, cStat, chNFSe, chDPS, nProt, motivo, nsNRec, erros.

    {
        "statusEnvio": "200",
        "statusConsulta": "200",
        "statusDownload": "200",
        "cStat": "100",
        "chNFSe": "31062002200000000000000000000000002926016153125604",
        "chDPS": "310620020000000000000000001000000000000029",
        "nProt": "",
        "motivo": "Autorizado o uso da NFS-e",
        "nsNRec": "313022",
        "erros": null
    }

Para ler o retorno, use as funções do módulo **modJSON**:

    Dim ret As Object
    Dim statusEnvio As String, statusConsulta As String, statusDownload As String
    Dim cStat As String, motivo As String, erros As String

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
                MsgBox motivo
                ' Testa se o download teve problemas
                If statusDownload <> "200" Then
                    MsgBox "Erro no Download"
                End If
            Else
                MsgBox motivo
            End If
        Else
            MsgBox motivo & vbCrLf & erros
        End If
    Else
        MsgBox motivo & vbCrLf & erros
    End If

`JsonStr` também aceita caminhos com ponto, por exemplo `JsonStr(ret, "erro.cStat")`.

##### Log:

Todas as requisições e respostas são gravadas em `<pasta do exe>\log\ddmmaaaa_NFSE.txt`. No download com sucesso só o status é gravado, para não encher o log com o PDF.

-----

Qualquer dúvida contate nossa equipe pelo WhatsApp ou/e HelpDesk. Obrigado pela atenção!!
