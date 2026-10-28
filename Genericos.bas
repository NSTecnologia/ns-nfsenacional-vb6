Attribute VB_Name = "Genericos"
'==============================================================================
' Genericos - Rotinas de apoio: log, arquivos, UTF-8 e Base64.
' Usa apenas APIs do Windows (kernel32, crypt32, shell32).
'==============================================================================
Option Explicit

Private Const CP_UTF8 As Long = 65001
Private Const CRYPT_STRING_BASE64 As Long = 1
Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Private Declare Function CryptStringToBinaryW Lib "crypt32" ( _
    ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, _
    ByVal pbBinary As Long, ByRef pcbBinary As Long, _
    ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

'------------------------------------------------------------------------------
' Arquivos do documento
'------------------------------------------------------------------------------
Public Sub salvarXML(ByVal xml As String, ByVal caminho As String, ByVal nome As String, _
                     Optional ByVal tpEvento As String = "", Optional ByVal nSeqEvento As String = "")
    gravarArquivoUtf8 caminho & tpEvento & nome & nSeqEvento & ".xml", xml
End Sub

Public Sub salvarJSON(ByVal json As String, ByVal caminho As String, ByVal nome As String, _
                      Optional ByVal tpEvento As String = "", Optional ByVal nSeqEvento As String = "")
    gravarArquivoUtf8 caminho & tpEvento & nome & nSeqEvento & ".json", json
End Sub

Public Sub salvarPDF(ByVal pdf As String, ByVal caminho As String, ByVal nome As String, _
                     Optional ByVal tpEvento As String = "", Optional ByVal nSeqEvento As String = "")
    Dim bytes() As Byte
    bytes = base64ParaBytes(pdf)
    gravarBytes caminho & tpEvento & nome & nSeqEvento & ".pdf", bytes
End Sub

'------------------------------------------------------------------------------
' Pasta do executável, sempre terminada em "\".
' Remove o prefixo "\\?\" que alguns ambientes (ex.: twinBASIC) devolvem.
'------------------------------------------------------------------------------
Public Function pastaApp() As String
    Dim p As String

    p = App.Path
    If Left$(p, 8) = "\\?\UNC\" Then
        p = "\\" & Mid$(p, 9)
    ElseIf Left$(p, 4) = "\\?\" Then
        p = Mid$(p, 5)
    End If

    If Right$(p, 1) <> "\" Then p = p & "\"
    pastaApp = p
End Function

Public Sub abrirArquivo(ByVal arquivo As String)
    ShellExecute 0, "open", arquivo, vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

'------------------------------------------------------------------------------
' Log diário em <pasta do exe>\log\ddmmaaaa_NFSE.txt
'------------------------------------------------------------------------------
Public Sub gravarLinhaLog(ByVal modelo As String, ByVal conteudo As String)
    Dim caminho As String
    Dim nomeArq As String
    Dim linha As String
    Dim bytes() As Byte
    Dim f As Integer

    On Error GoTo Fim

    caminho = pastaApp() & "log\"
    criarPasta caminho

    nomeArq = caminho & Format$(Now, "ddmmyyyy") & "_" & Trim$(Replace(modelo, ":", "")) & ".txt"
    linha = Format$(Now, "dd/mm/yyyy hh:nn") & " - " & conteudo & vbCrLf
    bytes = textoParaUtf8(linha)

    f = FreeFile
    Open nomeArq For Binary Access Write As #f
    Seek #f, LOF(f) + 1
    Put #f, , bytes
    Close #f

Fim:
End Sub

'------------------------------------------------------------------------------
' Cria a pasta (e as intermediárias). Retorna False se não conseguir.
'------------------------------------------------------------------------------
Public Function criarPasta(ByVal caminho As String) As Boolean
    Dim partes() As String
    Dim atual As String
    Dim i As Long

    On Error GoTo Falha

    If Right$(caminho, 1) = "\" Then caminho = Left$(caminho, Len(caminho) - 1)
    If Len(caminho) = 0 Then Exit Function

    partes = Split(caminho, "\")
    For i = 0 To UBound(partes)
        If i = 0 Then
            atual = partes(0)
        Else
            atual = atual & "\" & partes(i)
        End If

        ' Pula a unidade (C:) e o início de caminhos de rede (\\servidor\compart)
        If Len(partes(i)) > 0 And Right$(atual, 1) <> ":" And Not (Left$(caminho, 2) = "\\" And i < 4) Then
            If Len(Dir$(atual, vbDirectory)) = 0 Then MkDir atual
        End If
    Next

    criarPasta = True
    Exit Function

Falha:
    criarPasta = False
End Function

'------------------------------------------------------------------------------
' Leitura e gravação de arquivos em UTF-8
'------------------------------------------------------------------------------
Public Function lerArquivoUtf8(ByVal arquivo As String) As String
    Dim bytes() As Byte
    Dim f As Integer

    f = FreeFile
    Open arquivo For Binary Access Read As #f
    If LOF(f) > 0 Then
        ReDim bytes(0 To LOF(f) - 1)
        Get #f, , bytes
    End If
    Close #f

    lerArquivoUtf8 = utf8ParaTexto(bytes)
End Function

Public Sub gravarArquivoUtf8(ByVal arquivo As String, ByVal conteudo As String)
    Dim bytes() As Byte
    bytes = textoParaUtf8(conteudo)
    gravarBytes arquivo, bytes
End Sub

Public Sub gravarBytes(ByVal arquivo As String, ByRef bytes() As Byte)
    Dim f As Integer

    If Len(Dir$(arquivo)) > 0 Then Kill arquivo

    f = FreeFile
    Open arquivo For Binary Access Write As #f
    If tamanhoArray(bytes) > 0 Then Put #f, , bytes
    Close #f
End Sub

'------------------------------------------------------------------------------
' Conversões UTF-8 <-> texto do VB6 (Unicode)
'------------------------------------------------------------------------------
Public Function textoParaUtf8(ByVal texto As String) As Byte()
    Dim bytes() As Byte
    Dim n As Long

    If Len(texto) = 0 Then
        textoParaUtf8 = bytes
        Exit Function
    End If

    n = WideCharToMultiByte(CP_UTF8, 0, StrPtr(texto), Len(texto), 0, 0, 0, 0)
    ReDim bytes(0 To n - 1)
    WideCharToMultiByte CP_UTF8, 0, StrPtr(texto), Len(texto), VarPtr(bytes(0)), n, 0, 0

    textoParaUtf8 = bytes
End Function

Public Function utf8ParaTexto(ByRef bytes() As Byte) As String
    Dim total As Long
    Dim inicio As Long
    Dim n As Long
    Dim texto As String

    total = tamanhoArray(bytes)
    If total = 0 Then Exit Function

    inicio = LBound(bytes)

    ' Ignora o BOM UTF-8 (EF BB BF), se existir
    If total >= 3 Then
        If bytes(inicio) = &HEF And bytes(inicio + 1) = &HBB And bytes(inicio + 2) = &HBF Then
            inicio = inicio + 3
            total = total - 3
            If total = 0 Then Exit Function
        End If
    End If

    n = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytes(inicio)), total, 0, 0)
    texto = String$(n, vbNullChar)
    MultiByteToWideChar CP_UTF8, 0, VarPtr(bytes(inicio)), total, StrPtr(texto), n

    utf8ParaTexto = texto
End Function

'------------------------------------------------------------------------------
' Base64 -> bytes (via crypt32)
'------------------------------------------------------------------------------
Public Function base64ParaBytes(ByVal base64 As String) As Byte()
    Dim bytes() As Byte
    Dim n As Long

    If Len(base64) = 0 Then
        base64ParaBytes = bytes
        Exit Function
    End If

    If CryptStringToBinaryW(StrPtr(base64), Len(base64), CRYPT_STRING_BASE64, 0, n, 0, 0) = 0 Or n = 0 Then
        base64ParaBytes = bytes
        Exit Function
    End If

    ReDim bytes(0 To n - 1)
    CryptStringToBinaryW StrPtr(base64), Len(base64), CRYPT_STRING_BASE64, VarPtr(bytes(0)), n, 0, 0

    base64ParaBytes = bytes
End Function

' Quantidade de itens de um array de bytes (0 se não inicializado).
Public Function tamanhoArray(ByRef bytes() As Byte) As Long
    On Error GoTo Vazio
    tamanhoArray = UBound(bytes) - LBound(bytes) + 1
    Exit Function
Vazio:
    tamanhoArray = 0
End Function
