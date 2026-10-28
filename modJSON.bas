Attribute VB_Name = "modJSON"
'==============================================================================
' modJSON - Leitor/gerador de JSON simples, sem dependências externas.
'
'   Objeto JSON  -> Scripting.Dictionary
'   Array JSON   -> Collection
'   String       -> String
'   Número       -> String (mantém o texto original, ex.: "200", "4302.54")
'   true/false   -> Boolean
'   null         -> Null
'==============================================================================
Option Explicit

Private m_txt As String
Private m_pos As Long
Private m_len As Long

'------------------------------------------------------------------------------
' Converte um texto JSON em Dictionary/Collection.
' Retorna Nothing se o texto estiver vazio ou não for um JSON válido.
'------------------------------------------------------------------------------
Public Function JsonParse(ByVal texto As String) As Object
    Dim valor As Variant

    On Error GoTo Falha

    m_txt = texto
    m_pos = 1
    m_len = Len(texto)

    PularEspacos
    If m_pos > m_len Then Exit Function

    Atribuir valor, LerValor()
    If IsObject(valor) Then Set JsonParse = valor
    Exit Function

Falha:
    Set JsonParse = Nothing
End Function

'------------------------------------------------------------------------------
' Lê um valor texto pelo caminho (ex.: "erro.cStat").
' Retorna "" se o caminho não existir.
'------------------------------------------------------------------------------
Public Function JsonStr(ByVal raiz As Object, ByVal caminho As String) As String
    Dim valor As Variant

    If Not BuscarCaminho(raiz, caminho, valor) Then Exit Function
    If IsObject(valor) Then Exit Function
    If IsNull(valor) Or IsEmpty(valor) Then Exit Function

    JsonStr = CStr(valor)
End Function

'------------------------------------------------------------------------------
' Lê um objeto/array pelo caminho (ex.: "erros").
' Retorna Nothing se o caminho não existir ou não for objeto/array.
'------------------------------------------------------------------------------
Public Function JsonObj(ByVal raiz As Object, ByVal caminho As String) As Object
    Dim valor As Variant

    If Not BuscarCaminho(raiz, caminho, valor) Then Exit Function
    If IsObject(valor) Then Set JsonObj = valor
End Function

'------------------------------------------------------------------------------
' Converte Dictionary/Collection/valor simples em texto JSON.
'------------------------------------------------------------------------------
Public Function JsonStringify(ByVal valor As Variant) As String
    Dim partes As String
    Dim chave As Variant
    Dim item As Variant

    If IsObject(valor) Then
        If valor Is Nothing Then
            JsonStringify = "null"

        ElseIf TypeName(valor) = "Dictionary" Then
            For Each chave In valor.Keys
                If Len(partes) > 0 Then partes = partes & ","
                partes = partes & """" & JsonEscapar(CStr(chave)) & """:" & JsonStringify(valor.Item(chave))
            Next
            JsonStringify = "{" & partes & "}"

        ElseIf TypeName(valor) = "Collection" Then
            For Each item In valor
                If Len(partes) > 0 Then partes = partes & ","
                partes = partes & JsonStringify(item)
            Next
            JsonStringify = "[" & partes & "]"

        Else
            JsonStringify = "null"
        End If

    ElseIf IsNull(valor) Or IsEmpty(valor) Then
        JsonStringify = "null"

    ElseIf VarType(valor) = vbBoolean Then
        JsonStringify = IIf(valor, "true", "false")

    Else
        JsonStringify = """" & JsonEscapar(CStr(valor)) & """"
    End If
End Function

'------------------------------------------------------------------------------
' Monta um par "nome":"valor" já escapado.
'------------------------------------------------------------------------------
Public Function JsonPar(ByVal nome As String, ByVal valor As String) As String
    JsonPar = """" & JsonEscapar(nome) & """:""" & JsonEscapar(valor) & """"
End Function

'------------------------------------------------------------------------------
' Escapa um texto para uso dentro de uma string JSON.
'------------------------------------------------------------------------------
Public Function JsonEscapar(ByVal texto As String) As String
    Dim i As Long
    Dim c As Integer
    Dim saida As String

    For i = 1 To Len(texto)
        c = AscW(Mid$(texto, i, 1))
        Select Case c
            Case 34: saida = saida & "\"""
            Case 92: saida = saida & "\\"
            Case 8:  saida = saida & "\b"
            Case 12: saida = saida & "\f"
            Case 10: saida = saida & "\n"
            Case 13: saida = saida & "\r"
            Case 9:  saida = saida & "\t"
            Case 0 To 31
                saida = saida & "\u" & Right$("000" & Hex$(c), 4)
            Case Else
                saida = saida & ChrW$(c)
        End Select
    Next

    JsonEscapar = saida
End Function

'==============================================================================
' Rotinas internas
'==============================================================================

Private Function BuscarCaminho(ByVal raiz As Object, ByVal caminho As String, ByRef valor As Variant) As Boolean
    Dim partes() As String
    Dim i As Long
    Dim atual As Variant

    If raiz Is Nothing Then Exit Function

    Set atual = raiz
    partes = Split(caminho, ".")

    For i = 0 To UBound(partes)
        If Not IsObject(atual) Then Exit Function
        If atual Is Nothing Then Exit Function
        If TypeName(atual) <> "Dictionary" Then Exit Function
        If Not atual.Exists(partes(i)) Then Exit Function
        Atribuir atual, atual.Item(partes(i))
    Next

    Atribuir valor, atual
    BuscarCaminho = True
End Function

' Atribui um Variant que pode ser objeto (Set) ou valor simples.
Private Sub Atribuir(ByRef destino As Variant, ByVal origem As Variant)
    If IsObject(origem) Then
        Set destino = origem
    Else
        destino = origem
    End If
End Sub

Private Sub PularEspacos()
    Dim c As String
    Do While m_pos <= m_len
        c = Mid$(m_txt, m_pos, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then
            m_pos = m_pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function LerValor() As Variant
    Dim c As String

    PularEspacos
    If m_pos > m_len Then Err.Raise vbObjectError + 1, "modJSON", "Fim inesperado do JSON"

    c = Mid$(m_txt, m_pos, 1)
    Select Case c
        Case "{"
            Set LerValor = LerObjeto()
        Case "["
            Set LerValor = LerArray()
        Case """"
            LerValor = LerString()
        Case "t"
            LerPalavra "true"
            LerValor = True
        Case "f"
            LerPalavra "false"
            LerValor = False
        Case "n"
            LerPalavra "null"
            LerValor = Null
        Case Else
            LerValor = LerNumero()
    End Select
End Function

Private Function LerObjeto() As Object
    Dim dic As Object
    Dim chave As String
    Dim valor As Variant

    Set dic = CreateObject("Scripting.Dictionary")
    m_pos = m_pos + 1 ' {

    PularEspacos
    If Mid$(m_txt, m_pos, 1) = "}" Then
        m_pos = m_pos + 1
        Set LerObjeto = dic
        Exit Function
    End If

    Do
        PularEspacos
        If Mid$(m_txt, m_pos, 1) <> """" Then Err.Raise vbObjectError + 2, "modJSON", "Chave esperada na posição " & m_pos
        chave = LerString()

        PularEspacos
        If Mid$(m_txt, m_pos, 1) <> ":" Then Err.Raise vbObjectError + 3, "modJSON", "':' esperado na posição " & m_pos
        m_pos = m_pos + 1

        Atribuir valor, LerValor()
        If dic.Exists(chave) Then dic.Remove chave
        dic.Add chave, valor

        PularEspacos
        Select Case Mid$(m_txt, m_pos, 1)
            Case ","
                m_pos = m_pos + 1
            Case "}"
                m_pos = m_pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 4, "modJSON", "',' ou '}' esperado na posição " & m_pos
        End Select
    Loop

    Set LerObjeto = dic
End Function

Private Function LerArray() As Collection
    Dim col As Collection
    Dim valor As Variant

    Set col = New Collection
    m_pos = m_pos + 1 ' [

    PularEspacos
    If Mid$(m_txt, m_pos, 1) = "]" Then
        m_pos = m_pos + 1
        Set LerArray = col
        Exit Function
    End If

    Do
        Atribuir valor, LerValor()
        col.Add valor

        PularEspacos
        Select Case Mid$(m_txt, m_pos, 1)
            Case ","
                m_pos = m_pos + 1
            Case "]"
                m_pos = m_pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 5, "modJSON", "',' ou ']' esperado na posição " & m_pos
        End Select
    Loop

    Set LerArray = col
End Function

' Lê uma string JSON. Usa um buffer pré-alocado para ser rápido
' mesmo com textos grandes (XML e PDF em Base64).
Private Function LerString() As String
    Dim buf As String
    Dim usado As Long
    Dim inicio As Long
    Dim i As Long
    Dim c As Integer
    Dim e As String

    buf = Space$(256)
    m_pos = m_pos + 1 ' aspas de abertura
    inicio = m_pos
    i = m_pos

    Do While i <= m_len
        c = AscW(Mid$(m_txt, i, 1))

        If c = 34 Then ' fim da string
            Anexar buf, usado, Mid$(m_txt, inicio, i - inicio)
            m_pos = i + 1
            LerString = Left$(buf, usado)
            Exit Function

        ElseIf c = 92 Then ' barra invertida
            Anexar buf, usado, Mid$(m_txt, inicio, i - inicio)
            e = Mid$(m_txt, i + 1, 1)
            Select Case e
                Case """": Anexar buf, usado, """"
                Case "\":  Anexar buf, usado, "\"
                Case "/":  Anexar buf, usado, "/"
                Case "b":  Anexar buf, usado, vbBack
                Case "f":  Anexar buf, usado, vbFormFeed
                Case "n":  Anexar buf, usado, vbLf
                Case "r":  Anexar buf, usado, vbCr
                Case "t":  Anexar buf, usado, vbTab
                Case "u"
                    Anexar buf, usado, ChrW$(CLng("&H" & Mid$(m_txt, i + 2, 4)))
                    i = i + 4
                Case Else
                    Err.Raise vbObjectError + 6, "modJSON", "Escape inválido na posição " & i
            End Select
            i = i + 2
            inicio = i
        Else
            i = i + 1
        End If
    Loop

    Err.Raise vbObjectError + 7, "modJSON", "String sem fechamento"
End Function

Private Sub Anexar(ByRef buf As String, ByRef usado As Long, ByVal trecho As String)
    Dim n As Long
    Dim cresc As Long

    n = Len(trecho)
    If n = 0 Then Exit Sub

    If usado + n > Len(buf) Then
        cresc = Len(buf)
        If cresc < n Then cresc = n
        buf = buf & Space$(cresc)
    End If

    Mid$(buf, usado + 1, n) = trecho
    usado = usado + n
End Sub

Private Function LerNumero() As String
    Dim inicio As Long
    Dim c As String

    inicio = m_pos
    Do While m_pos <= m_len
        c = Mid$(m_txt, m_pos, 1)
        If InStr(1, "+-0123456789.eE", c, vbBinaryCompare) > 0 Then
            m_pos = m_pos + 1
        Else
            Exit Do
        End If
    Loop

    If m_pos = inicio Then Err.Raise vbObjectError + 8, "modJSON", "Valor inválido na posição " & m_pos
    LerNumero = Mid$(m_txt, inicio, m_pos - inicio)
End Function

Private Sub LerPalavra(ByVal palavra As String)
    If Mid$(m_txt, m_pos, Len(palavra)) <> palavra Then
        Err.Raise vbObjectError + 9, "modJSON", "Valor inválido na posição " & m_pos
    End If
    m_pos = m_pos + Len(palavra)
End Sub
