Attribute VB_Name = "modDREExport"
'==================================================================
' modDREExport  -  Exportador DRE Interativa  (Mac-compatible)
'------------------------------------------------------------------
' Lê as abas "Entradas" e "Saídas" diretamente neste workbook,
' filtra pelo período configurado, valida e injeta JSON no HTML.
'
' MUDANÇAS em relação à versão anterior:
'   - Dados de entradas e saídas ficam em abas deste próprio arquivo
'     (não há mais caminhos entradas_path / saidas_path na Config).
'   - Sem CreateObject("Scripting.Dictionary") nem ArrayList:
'     compatível 100% com Excel para Mac.
'
' Aba Config (coluna A = chave, coluna B em diante = valor):
'   html_template     caminho do dre_interativa.html de origem
'   html_output       caminho do HTML gerado
'   periodo_mes_ini   mês de início  (1-12)
'   periodo_ano_ini   ano  de início (ex.: 2025)
'   periodo_mes_fim   mês  de fim    (1-12)   — opcional
'   periodo_ano_fim   ano  de fim    (ex.: 2025) — opcional
'   entradas_abas     nome(s) da(s) aba(s) de Entradas  (padrão: "Entradas")
'   saidas_abas       nome(s) da(s) aba(s) de Saídas    (padrão: "Saídas")
'   json_output       caminho do dre_data.json gerado (opcional)
'                     Se preenchido, exporta também um JSON puro para upload manual
'                     no painel da Vercel (env var DRE_DATA_JSON).
'
'   Se periodo_mes_fim / periodo_ano_fim forem omitidos, o fim do período é
'   determinado automaticamente pela data mais recente encontrada nos dados.
'==================================================================

Option Explicit

Private Const CFG_HTML_IN  As String = "html_template"
Private Const CFG_HTML_OUT As String = "html_output"
Private Const CFG_MES_INI  As String = "periodo_mes_ini"
Private Const CFG_ANO_INI  As String = "periodo_ano_ini"
Private Const CFG_MES_FIM  As String = "periodo_mes_fim"
Private Const CFG_ANO_FIM  As String = "periodo_ano_fim"
Private Const CFG_ENT_ABAS As String = "entradas_abas"
Private Const CFG_SAI_ABAS As String = "saidas_abas"
Private Const CFG_API_KEY  As String = "chave_api"
Private Const CFG_JSON_OUT As String = "json_output"

Private Const MARK_START As String = "/*{{DRE_DATA_START}}*/"
Private Const MARK_END   As String = "/*{{DRE_DATA_END}}*/"

' Colunas de Entradas (1-based, linha 1 = cabeçalho)
Private Const E_EMPRESA    As Long = 1   ' A
Private Const E_PACOTE     As Long = 2   ' B
Private Const E_NF         As Long = 3   ' C
Private Const E_CLIENTE    As Long = 4   ' D
Private Const E_VLR_FAT    As Long = 5   ' E
Private Const E_VLR_REC    As Long = 6   ' F
Private Const E_OBJETO     As Long = 7   ' G
Private Const E_DT_EMISSAO As Long = 8   ' H  <- filtro de período
Private Const E_DT_VENC    As Long = 9   ' I
Private Const E_STATUS     As Long = 10  ' J
Private Const E_COLUNA1    As Long = 11  ' K  (auxiliar — ignorado)
Private Const E_DT_RECEB   As Long = 12  ' L
Private Const E_VERT1      As Long = 13  ' M
Private Const E_PCT_V1     As Long = 14  ' N
Private Const E_VERT2      As Long = 15  ' O
Private Const E_PCT_V2     As Long = 16  ' P
Private Const E_OBS        As Long = 17  ' Q  (auxiliar — ignorado)
Private Const E_COBRANCA   As Long = 18  ' R  (auxiliar — ignorado)
Private Const E_SEGMENTO   As Long = 19  ' S

' Colunas de Saídas (1-based, linha 1 = cabeçalho)
Private Const S_DT_VENC    As Long = 1   ' A  <- filtro de período
Private Const S_DT_PAG     As Long = 2   ' B
Private Const S_FORNECEDOR As Long = 3   ' C
Private Const S_DOC        As Long = 4   ' D
Private Const S_DESCRICAO  As Long = 5   ' E
Private Const S_CC         As Long = 6   ' F
Private Const S_GRUPO      As Long = 7   ' G
Private Const S_TIPO_GASTO As Long = 8   ' H
Private Const S_VALOR      As Long = 9   ' I
Private Const S_BANCO      As Long = 10  ' J
Private Const S_OBS        As Long = 11  ' K


'==================================================================
' PONTO DE ENTRADA
'==================================================================
Public Sub ExportarDRE()
    On Error GoTo fail

    Dim t0 As Double: t0 = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    '--- 1. Ler configurações ----------------------------------------
    Dim htmlIn As String, htmlOut As String
    Dim pAnoIni As Long, pMesIni As Long
    Dim pAnoFim As Long, pMesFim As Long
    Dim fimAutomatico As Boolean
    Dim entAbas() As String, saiAbas() As String

    LerConfig htmlIn, htmlOut, pAnoIni, pMesIni, pAnoFim, pMesFim, _
              fimAutomatico, entAbas, saiAbas

    If pMesIni < 1 Or pMesIni > 12 Then
        Err.Raise vbObjectError + 2102, , _
            "período_mes_ini inválido na Config: " & pMesIni & ". Deve estar entre 1 e 12."
    End If
    If Not fimAutomatico Then
        If pMesFim < 1 Or pMesFim > 12 Then
            Err.Raise vbObjectError + 2102, , _
                "periodo_mes_fim inválido na Config: " & pMesFim & ". Deve estar entre 1 e 12."
        End If
        If pAnoIni * 100 + pMesIni > pAnoFim * 100 + pMesFim Then
            Err.Raise vbObjectError + 2102, , _
                "Período inválido: início (" & pMesIni & "/" & pAnoIni & ") " & _
                "é posterior ao fim (" & pMesFim & "/" & pAnoFim & ")."
        End If
    End If

    '--- 2. Validar que as abas existem neste workbook ---------------
    ValidarAbas entAbas, "Entradas"
    ValidarAbas saiAbas, "Saídas"

    '--- 3. Ler mapeamento Grupo -> Rubrica --------------------------
    Dim mapaGrupos() As String, mapaRubIds() As String
    Dim mapaRubs() As String,   mapaOrds() As Long
    Dim mapaCount As Long
    LerMapeamento mapaGrupos, mapaRubIds, mapaRubs, mapaOrds, mapaCount

    If mapaCount = 0 Then
        Err.Raise vbObjectError + 2001, , "Aba 'Mapeamento' está vazia ou não foi encontrada."
    End If

    '--- 4. Determinar fim automático (se não configurado) -----------
    If fimAutomatico Then
        AcharUltimaData entAbas, saiAbas, pAnoFim, pMesFim
    End If

    '--- 5. Validar INVESTIGAR no período ----------------------------
    ValidarInvestigar saiAbas, pAnoIni, pMesIni, pAnoFim, pMesFim

    '--- 6. Validar que todos os Grupos têm mapeamento ---------------
    ValidarGruposMapeados saiAbas, mapaGrupos, mapaCount, pAnoIni, pMesIni, pAnoFim, pMesFim

    '--- 7. Construir JSON -------------------------------------------
    Dim apiKey As String: apiKey = LerApiKey()
    Dim jsonEnt As String: jsonEnt = LerEntradasJSON(entAbas, pAnoIni, pMesIni, pAnoFim, pMesFim)
    Dim jsonSai As String: jsonSai = LerSaidasJSON(saiAbas, pAnoIni, pMesIni, pAnoFim, pMesFim)
    Dim jsonMap As String: jsonMap = MapeamentoJSON(mapaGrupos, mapaRubIds, mapaRubs, mapaOrds, mapaCount)

    '--- 8. Montar payload completo ----------------------------------
    Dim payload As String
    payload = "{" & vbCrLf & _
        "  ""meta"": {" & vbCrLf & _
        "    ""generatedAt"": """ & Format(Now, "yyyy-mm-dd hh:nn:ss") & """," & vbCrLf & _
        "    ""regime"": ""competencia""," & vbCrLf & _
        "    ""apiKey"": """ & apiKey & """," & vbCrLf & _
        "    ""periodoAnoIni"": " & pAnoIni & "," & vbCrLf & _
        "    ""periodoMesIni"": " & pMesIni & "," & vbCrLf & _
        "    ""periodoAnoFim"": " & pAnoFim & "," & vbCrLf & _
        "    ""periodoMesFim"": " & pMesFim & vbCrLf & _
        "  }," & vbCrLf & _
        "  ""mapeamento"": " & jsonMap & "," & vbCrLf & _
        "  ""entradas"": " & jsonEnt & "," & vbCrLf & _
        "  ""saidas"": " & jsonSai & vbCrLf & _
        "}"

    '--- 9. Injetar no HTML ------------------------------------------
    InjetarNoHTML htmlIn, htmlOut, payload

    '--- 10. Exportar JSON separado (opcional) -----------------------
    Dim jsonOut As String: jsonOut = CfgStrOpt(CFG_JSON_OUT)
    If Len(jsonOut) > 0 Then ExportarJSON payload, jsonOut

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    Dim msgFim As String
    If fimAutomatico Then
        msgFim = Format(DateSerial(pAnoFim, pMesFim, 1), "mmm/yy") & " (automático)"
    Else
        msgFim = Format(DateSerial(pAnoFim, pMesFim, 1), "mmm/yy")
    End If

    MsgBox "DRE exportada com sucesso!" & vbCrLf & vbCrLf & _
           "Período : " & Format(DateSerial(pAnoIni, pMesIni, 1), "mmm/yy") & _
           " a " & msgFim & vbCrLf & _
           "Arquivo : " & htmlOut & vbCrLf & _
           "Tempo   : " & Format(Timer - t0, "0.00") & "s", _
           vbInformation, "DRE Interativa"
    Exit Sub

fail:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    ' Captura ANTES de qualquer On Error que zeraria Err
    Dim errNum As Long:    errNum  = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    MsgBox "Falha na exportação:" & vbCrLf & vbCrLf & errDesc, _
           vbCritical, "DRE Interativa — erro " & errNum
End Sub


'==================================================================
' CONFIG  (aba "Config")
'   Coluna A = chave | Coluna B [C D …] = valor(es)
'==================================================================
Private Sub LerConfig(htmlIn As String, htmlOut As String, _
                      pAnoIni As Long, pMesIni As Long, _
                      pAnoFim As Long, pMesFim As Long, _
                      fimAutomatico As Boolean, _
                      entAbas() As String, saiAbas() As String)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2100, , "Aba 'Config' não encontrada neste workbook."
    End If

    htmlIn  = CfgStr(ws, CFG_HTML_IN)
    htmlOut = CfgStr(ws, CFG_HTML_OUT)
    pMesIni = CfgInt(ws, CFG_MES_INI)
    pAnoIni = CfgInt(ws, CFG_ANO_INI)

    Dim mesFimOpt As Long, anoFimOpt As Long
    mesFimOpt = CfgIntOpt(ws, CFG_MES_FIM)
    anoFimOpt = CfgIntOpt(ws, CFG_ANO_FIM)

    fimAutomatico = (mesFimOpt = 0 Or anoFimOpt = 0)
    If Not fimAutomatico Then
        pMesFim = mesFimOpt
        pAnoFim = anoFimOpt
    End If

    entAbas = CfgLst(ws, CFG_ENT_ABAS, "Entradas")
    saiAbas = CfgLst(ws, CFG_SAI_ABAS, "Saídas")
End Sub

' Retorna a linha (1-based) onde a chave está, ou 0 se não encontrada
Private Function CfgRow(ws As Worksheet, key As String) As Long
    CfgRow = 0
    Dim r As Long: r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), key, vbTextCompare) = 0 Then
            CfgRow = r: Exit Function
        End If
        r = r + 1
        If r > 500 Then Exit Do
    Loop
End Function

Private Function CfgStr(ws As Worksheet, key As String) As String
    Dim r As Long: r = CfgRow(ws, key)
    If r = 0 Then
        Err.Raise vbObjectError + 2101, , "Chave '" & key & "' não encontrada na aba Config."
    End If
    CfgStr = Trim$(CStr(ws.Cells(r, 2).Value))
    If Len(CfgStr) = 0 Then
        Err.Raise vbObjectError + 2101, , "Chave '" & key & "' está vazia na aba Config."
    End If
End Function

Private Function CfgInt(ws As Worksheet, key As String) As Long
    Dim s As String: s = CfgStr(ws, key)
    If Not IsNumeric(s) Then
        Err.Raise vbObjectError + 2103, , "Config '" & key & "' não é numérico: """ & s & """"
    End If
    CfgInt = CLng(CDbl(s))
End Function

' Retorna 0 se a chave estiver ausente ou vazia (campo opcional).
Private Function CfgIntOpt(ws As Worksheet, key As String) As Long
    Dim r As Long: r = CfgRow(ws, key)
    If r = 0 Then CfgIntOpt = 0: Exit Function
    Dim s As String: s = Trim$(CStr(ws.Cells(r, 2).Value))
    If Len(s) = 0 Or Not IsNumeric(s) Then CfgIntOpt = 0: Exit Function
    CfgIntOpt = CLng(CDbl(s))
End Function

' Lê lista de valores (colunas B, C, D…). Se chave ausente, usa defaultVal.
Private Function CfgLst(ws As Worksheet, key As String, defaultVal As String) As String()
    Dim result() As String
    Dim r As Long: r = CfgRow(ws, key)

    If r = 0 Then
        ReDim result(0): result(0) = defaultVal
        CfgLst = result: Exit Function
    End If

    Dim c As Long, n As Long: n = 0
    For c = 2 To 200
        If Len(Trim$(CStr(ws.Cells(r, c).Value))) = 0 Then Exit For
        n = n + 1
    Next c

    If n = 0 Then
        ReDim result(0): result(0) = defaultVal
        CfgLst = result: Exit Function
    End If

    ReDim result(n - 1)
    For c = 0 To n - 1
        result(c) = Trim$(CStr(ws.Cells(r, c + 2).Value))
    Next c
    CfgLst = result
End Function


Private Function LerApiKey() As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim r As Long: r = CfgRow(ws, CFG_API_KEY)
    If r = 0 Then Exit Function
    LerApiKey = Trim$(CStr(ws.Cells(r, 2).Value))
End Function

' Retorna "" se a chave estiver ausente ou vazia (campo de string opcional).
Private Function CfgStrOpt(key As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    Dim r As Long: r = CfgRow(ws, key)
    If r = 0 Then Exit Function
    CfgStrOpt = Trim$(CStr(ws.Cells(r, 2).Value))
End Function

' Grava o payload JSON em um arquivo de texto separado (dre_data.json).
Private Sub ExportarJSON(ByVal payload As String, ByVal caminho As String)
    Dim f As Integer: f = FreeFile()
    Open caminho For Output As #f
    Print #f, payload
    Close #f
End Sub


'==================================================================
' MAPEAMENTO  (aba "Mapeamento")
'   A=Grupo | B=RubricaId | C=Rubrica | D=Ordem
'==================================================================
Private Sub LerMapeamento(grupos() As String, rubIds() As String, _
                           rubs() As String, ords() As Long, n As Long)
    n = 0
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Mapeamento")
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2200, , "Aba 'Mapeamento' não encontrada."
    End If

    ' Contar linhas
    Dim r As Long: r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        n = n + 1: r = r + 1
        If r > 10000 Then Exit Do
    Loop
    If n = 0 Then Exit Sub

    ReDim grupos(n - 1): ReDim rubIds(n - 1)
    ReDim rubs(n - 1):   ReDim ords(n - 1)

    Dim i As Long
    For i = 0 To n - 1
        r = i + 2
        grupos(i) = Trim$(CStr(ws.Cells(r, 1).Value))
        rubIds(i) = Trim$(CStr(ws.Cells(r, 2).Value))
        rubs(i)   = Trim$(CStr(ws.Cells(r, 3).Value))
        Dim ordStr As String: ordStr = Trim$(CStr(ws.Cells(r, 4).Value))
        ords(i) = IIf(Len(ordStr) > 0 And IsNumeric(ordStr), CLng(ordStr), 99)

        If Len(rubIds(i)) = 0 Or Len(rubs(i)) = 0 Then
            Err.Raise vbObjectError + 2201, , _
                "Linha " & r & " da aba Mapeamento com RubricaId ou Rubrica vazio."
        End If
    Next i
End Sub

' Retorna índice (0-based) de um Grupo no array, ou -1 se não achado
Private Function MapaIdx(grupos() As String, n As Long, grupo As String) As Long
    MapaIdx = -1
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(grupos(i), grupo, vbTextCompare) = 0 Then
            MapaIdx = i: Exit Function
        End If
    Next i
End Function


'==================================================================
' VALIDAÇÕES
'==================================================================
Private Sub ValidarAbas(abas() As String, contexto As String)
    Dim i As Long
    For i = 0 To UBound(abas)
        Dim ws As Worksheet, found As Boolean: found = False
        For Each ws In ThisWorkbook.Worksheets
            If StrComp(ws.Name, abas(i), vbTextCompare) = 0 Then
                found = True: Exit For
            End If
        Next ws
        If Not found Then
            Err.Raise vbObjectError + 2400, , _
                "Aba '" & abas(i) & "' não encontrada neste workbook " & _
                "(esperada como aba de " & contexto & ")."
        End If
    Next i
End Sub

Private Sub ValidarInvestigar(abas() As String, pAnoIni As Long, pMesIni As Long, _
                              pAnoFim As Long, pMesFim As Long)
    Dim totInv As Long, sample As String
    Dim ai As Long
    For ai = 0 To UBound(abas)
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(abas(ai))
        Dim ultL As Long: ultL = ws.Cells(ws.Rows.Count, S_DESCRICAO).End(xlUp).Row
        If ultL < 2 Then GoTo proxAba

        Dim r As Long, g As String, tg As String
        For r = 2 To ultL
            If EstaNoPeriodo(ws.Cells(r, S_DT_VENC).Value, pAnoIni, pMesIni, pAnoFim, pMesFim) Then
                g  = UCase$(SafeStr(ws.Cells(r, S_GRUPO).Value))
                tg = UCase$(SafeStr(ws.Cells(r, S_TIPO_GASTO).Value))
                If g = "INVESTIGAR" Or tg = "INVESTIGAR" Then
                    totInv = totInv + 1
                    If totInv <= 5 Then
                        sample = sample & vbCrLf & _
                                 "  [" & ws.Name & " L" & r & "] " & _
                                 ws.Cells(r, S_DESCRICAO).Value & " / " & _
                                 ws.Cells(r, S_CC).Value & " (" & _
                                 Format(ws.Cells(r, S_VALOR).Value, "#,##0.00") & ")"
                    End If
                End If
            End If
        Next r
proxAba:
    Next ai

    If totInv > 0 Then
        Err.Raise vbObjectError + 2300, , _
            "Encontrados " & totInv & " lançamentos em INVESTIGAR no período. " & _
            "Classifique-os antes de exportar." & vbCrLf & _
            "Primeiras ocorrências:" & sample & _
            IIf(totInv > 5, vbCrLf & "  ... (+" & (totInv - 5) & " outros)", "")
    End If
End Sub

Private Sub ValidarGruposMapeados(abas() As String, mapaGrupos() As String, _
                                   mapaCount As Long, pAnoIni As Long, _
                                   pMesIni As Long, pAnoFim As Long, pMesFim As Long)
    ' Arrays para grupos não mapeados (máx 1000 grupos distintos)
    Dim faltGrupos(999) As String
    Dim faltConts(999)  As Long
    Dim faltN As Long: faltN = 0

    Dim ai As Long
    For ai = 0 To UBound(abas)
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(abas(ai))
        Dim ultL As Long: ultL = ws.Cells(ws.Rows.Count, S_DESCRICAO).End(xlUp).Row
        If ultL < 2 Then GoTo prox

        Dim r As Long, g As String
        For r = 2 To ultL
            If EstaNoPeriodo(ws.Cells(r, S_DT_VENC).Value, pAnoIni, pMesIni, pAnoFim, pMesFim) Then
                g = SafeStr(ws.Cells(r, S_GRUPO).Value)
                If Len(g) > 0 And Not EhMovFinanceira(g) Then
                    If MapaIdx(mapaGrupos, mapaCount, g) = -1 Then
                        Dim fi As Long, found As Boolean: found = False
                        For fi = 0 To faltN - 1
                            If StrComp(faltGrupos(fi), g, vbTextCompare) = 0 Then
                                faltConts(fi) = faltConts(fi) + 1
                                found = True: Exit For
                            End If
                        Next fi
                        If Not found And faltN < 1000 Then
                            faltGrupos(faltN) = g
                            faltConts(faltN) = 1
                            faltN = faltN + 1
                        End If
                    End If
                End If
            End If
        Next r
prox:
    Next ai

    If faltN > 0 Then
        Dim msg As String, k As Long
        For k = 0 To faltN - 1
            msg = msg & vbCrLf & "  '" & faltGrupos(k) & "' (" & faltConts(k) & " lanç.)"
        Next k
        Err.Raise vbObjectError + 2301, , _
            "Grupos sem mapeamento — adicione-os na aba Mapeamento:" & msg
    End If
End Sub


'==================================================================
' LEITURA -> JSON ARRAY
'==================================================================
Private Function LerEntradasJSON(abas() As String, _
                                  pAnoIni As Long, pMesIni As Long, _
                                  pAnoFim As Long, pMesFim As Long) As String
    Dim json As String: json = "["
    Dim first As Boolean: first = True

    Dim ai As Long
    For ai = 0 To UBound(abas)
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(abas(ai))
        Dim ultL As Long: ultL = ws.Cells(ws.Rows.Count, E_CLIENTE).End(xlUp).Row
        If ultL < 2 Then GoTo proxAbaE

        ' +1 garante array 2D mesmo com 1 única linha de dados
        Dim dados As Variant
        dados = ws.Range(ws.Cells(2, 1), ws.Cells(ultL + 1, E_SEGMENTO)).Value

        Dim r As Long
        For r = 1 To UBound(dados, 1)
            If Len(Trim$(CStr(dados(r, E_CLIENTE)))) = 0 And _
               Len(Trim$(CStr(dados(r, E_EMPRESA)))) = 0 Then GoTo proxLinhaE

            If Not EstaNoPeriodo(dados(r, E_DT_EMISSAO), pAnoIni, pMesIni, pAnoFim, pMesFim) Then GoTo proxLinhaE

            If Not first Then json = json & ","
            first = False

            json = json & vbCrLf & "    {" & _
                """empresa"":"""   & JsonEsc(CStr(dados(r, E_EMPRESA)))   & """," & _
                """pacote"":"""    & JsonEsc(CStr(dados(r, E_PACOTE)))    & """," & _
                """nf"":"""        & JsonEsc(CStr(dados(r, E_NF)))        & """," & _
                """cliente"":"""   & JsonEsc(CStr(dados(r, E_CLIENTE)))   & """," & _
                """vlrFat"":"      & NumJson(dados(r, E_VLR_FAT))         & ","  & _
                """vlrRec"":"      & NumJson(dados(r, E_VLR_REC))         & ","  & _
                """objeto"":"""    & JsonEsc(CStr(dados(r, E_OBJETO)))    & """," & _
                """dtEmissao"":"   & DateJson(dados(r, E_DT_EMISSAO))     & ","  & _
                """dtVenc"":"      & DateJson(dados(r, E_DT_VENC))        & ","  & _
                """status"":"""    & JsonEsc(CStr(dados(r, E_STATUS)))    & """," & _
                """dtReceb"":"     & DateJson(dados(r, E_DT_RECEB))       & ","  & _
                """v1"":"""        & JsonEsc(CStr(dados(r, E_VERT1)))     & """," & _
                """pctV1"":"       & NumJson(dados(r, E_PCT_V1))          & ","  & _
                """v2"":"""        & JsonEsc(CStr(dados(r, E_VERT2)))     & """," & _
                """pctV2"":"       & NumJson(dados(r, E_PCT_V2))          & ","  & _
                """segmento"":"""  & JsonEsc(CStr(dados(r, E_SEGMENTO)))  & """," & _
                """_aba"":"""      & JsonEsc(ws.Name)                      & """}"
proxLinhaE:
        Next r
proxAbaE:
    Next ai

    LerEntradasJSON = json & vbCrLf & "  ]"
End Function

Private Function LerSaidasJSON(abas() As String, _
                                pAnoIni As Long, pMesIni As Long, _
                                pAnoFim As Long, pMesFim As Long) As String
    Dim json As String: json = "["
    Dim first As Boolean: first = True

    Dim ai As Long
    For ai = 0 To UBound(abas)
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(abas(ai))
        Dim ultL As Long: ultL = ws.Cells(ws.Rows.Count, S_DESCRICAO).End(xlUp).Row
        If ultL < 2 Then GoTo proxAbaS

        ' +1 garante array 2D mesmo com 1 única linha de dados
        Dim dados As Variant
        dados = ws.Range(ws.Cells(2, 1), ws.Cells(ultL + 1, S_OBS)).Value

        Dim r As Long
        For r = 1 To UBound(dados, 1)
            If Len(Trim$(CStr(dados(r, S_DESCRICAO)))) = 0 And _
               Len(Trim$(CStr(dados(r, S_CC))))         = 0 And _
               Len(Trim$(CStr(dados(r, S_VALOR))))      = 0 Then GoTo proxLinhaS

            If Not EstaNoPeriodo(dados(r, S_DT_VENC), pAnoIni, pMesIni, pAnoFim, pMesFim) Then GoTo proxLinhaS

            If EhMovFinanceira(SafeStr(dados(r, S_GRUPO))) Then GoTo proxLinhaS

            If Not first Then json = json & ","
            first = False

            json = json & vbCrLf & "    {" & _
                """dtVenc"":"       & DateJson(dados(r, S_DT_VENC))        & ","  & _
                """dtPag"":"        & DateJson(dados(r, S_DT_PAG))         & ","  & _
                """fornecedor"":""" & JsonEsc(CStr(dados(r, S_FORNECEDOR)))& """," & _
                """doc"":"""        & JsonEsc(CStr(dados(r, S_DOC)))       & """," & _
                """descricao"":"""  & JsonEsc(CStr(dados(r, S_DESCRICAO))) & """," & _
                """cc"":"""         & JsonEsc(CStr(dados(r, S_CC)))        & """," & _
                """grupo"":"""      & JsonEsc(CStr(dados(r, S_GRUPO)))     & """," & _
                """tipoGasto"":"""  & JsonEsc(CStr(dados(r, S_TIPO_GASTO)))& """," & _
                """valor"":"        & NumJson(dados(r, S_VALOR))           & ","  & _
                """banco"":"""      & JsonEsc(CStr(dados(r, S_BANCO)))     & """," & _
                """obs"":"""        & JsonEsc(CStr(dados(r, S_OBS)))       & """," & _
                """_aba"":"""       & JsonEsc(ws.Name)                      & """}"
proxLinhaS:
        Next r
proxAbaS:
    Next ai

    LerSaidasJSON = json & vbCrLf & "  ]"
End Function

Private Function MapeamentoJSON(grupos() As String, rubIds() As String, _
                                 rubs() As String, ords() As Long, n As Long) As String
    Dim json As String: json = "["
    Dim i As Long
    For i = 0 To n - 1
        If i > 0 Then json = json & ","
        json = json & vbCrLf & "    {" & _
            """grupo"":"""     & JsonEsc(grupos(i)) & """," & _
            """rubricaId"":""" & JsonEsc(rubIds(i)) & """," & _
            """rubrica"":"""   & JsonEsc(rubs(i))   & """," & _
            """ordem"":"       & ords(i)             & "}"
    Next i
    MapeamentoJSON = json & vbCrLf & "  ]"
End Function


'==================================================================
' INJEÇÃO NO HTML
'==================================================================
Private Sub InjetarNoHTML(tplPath As String, outPath As String, payload As String)
    Dim html As String: html = LerArquivoTexto(tplPath)

    Dim p1 As Long, p2 As Long
    p1 = InStr(1, html, MARK_START, vbBinaryCompare)
    p2 = InStr(1, html, MARK_END, vbBinaryCompare)

    If p1 = 0 Or p2 = 0 Or p2 <= p1 Then
        Err.Raise vbObjectError + 2401, , _
            "Marcadores " & MARK_START & " / " & MARK_END & _
            " não encontrados no template:" & vbCrLf & tplPath & vbCrLf & _
            "Use a versão atual do dre_interativa.html."
    End If

    Dim antes As String, depois As String
    antes  = Left$(html, p1 + Len(MARK_START) - 1)
    depois = Mid$(html, p2)

    GravarArquivoTexto outPath, antes & vbCrLf & payload & vbCrLf & depois
End Sub


'==================================================================
' HELPERS — I/O, datas, JSON
'==================================================================
Private Function LerArquivoTexto(ByVal path As String) As String
    If Len(Dir(path)) = 0 Then
        Err.Raise vbObjectError + 2501, , "Arquivo não encontrado: " & path
    End If
    Dim iFile As Integer: iFile = FreeFile
    Open path For Input As #iFile
        LerArquivoTexto = Input$(LOF(iFile), iFile)
    Close #iFile
End Function

Private Sub GravarArquivoTexto(ByVal path As String, ByVal conteudo As String)
    Dim iFile As Integer: iFile = FreeFile
    Open path For Output As #iFile
        Print #iFile, conteudo;
    Close #iFile
End Sub

Private Function EstaNoPeriodo(v As Variant, pAnoIni As Long, pMesIni As Long, _
                               pAnoFim As Long, pMesFim As Long) As Boolean
    EstaNoPeriodo = False
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    Dim d As Date
    On Error Resume Next
    d = CDate(v)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    On Error GoTo 0
    Dim ym As Long: ym = Year(d) * 100 + Month(d)
    If ym < pAnoIni * 100 + pMesIni Then Exit Function
    If ym > pAnoFim * 100 + pMesFim Then Exit Function
    EstaNoPeriodo = True
End Function

' Varre todas as abas de entradas (dtEmissao) e saídas (dtVenc)
' e retorna o ano/mês da data mais recente encontrada.
Private Sub AcharUltimaData(entAbas() As String, saiAbas() As String, _
                             pAnoFim As Long, pMesFim As Long)
    Dim maxYM As Long: maxYM = 0
    Dim ai As Long, r As Long, ultL As Long
    Dim ws As Worksheet
    Dim v As Variant, d As Date, ym As Long

    For ai = 0 To UBound(entAbas)
        Set ws = ThisWorkbook.Sheets(entAbas(ai))
        ultL = ws.Cells(ws.Rows.Count, E_DT_EMISSAO).End(xlUp).Row
        For r = 2 To ultL
            v = ws.Cells(r, E_DT_EMISSAO).Value
            If Not (IsEmpty(v) Or IsNull(v)) Then
                On Error Resume Next
                d = CDate(v)
                If Err.Number = 0 Then
                    ym = Year(d) * 100 + Month(d)
                    If ym > maxYM Then maxYM = ym
                End If
                Err.Clear
                On Error GoTo 0
            End If
        Next r
    Next ai

    For ai = 0 To UBound(saiAbas)
        Set ws = ThisWorkbook.Sheets(saiAbas(ai))
        ultL = ws.Cells(ws.Rows.Count, S_DT_VENC).End(xlUp).Row
        For r = 2 To ultL
            v = ws.Cells(r, S_DT_VENC).Value
            If Not (IsEmpty(v) Or IsNull(v)) Then
                On Error Resume Next
                d = CDate(v)
                If Err.Number = 0 Then
                    ym = Year(d) * 100 + Month(d)
                    If ym > maxYM Then maxYM = ym
                End If
                Err.Clear
                On Error GoTo 0
            End If
        Next r
    Next ai

    If maxYM = 0 Then
        Err.Raise vbObjectError + 2104, , _
            "periodo_mes_fim / periodo_ano_fim não configurados e não foi possível " & _
            "determinar a data mais recente: nenhuma data válida encontrada nos dados."
    End If

    pAnoFim = maxYM \ 100
    pMesFim = maxYM Mod 100
End Sub

Private Function JsonEsc(ByVal s As String) As String
    If Len(s) = 0 Then JsonEsc = "": Exit Function
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    ' Escapa não-ASCII como \uXXXX: evita problemas de encoding MacRoman vs UTF-8
    Dim result As String, i As Long, cp As Long
    For i = 1 To Len(s)
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536  ' AscW retorna Integer sinalizado
        If cp > 127 Then
            result = result & "\u" & Right$("0000" & Hex$(cp), 4)
        Else
            result = result & Mid$(s, i, 1)
        End If
    Next i
    JsonEsc = result
End Function

Private Function NumJson(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then NumJson = "0": Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then NumJson = "0": Exit Function
    End If
    If Not IsNumeric(v) Then NumJson = "0": Exit Function
    NumJson = Replace(CStr(CDbl(v)), ",", ".")
End Function

Private Function DateJson(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then DateJson = "null": Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then DateJson = "null": Exit Function
    End If
    Dim d As Date
    On Error Resume Next
    d = CDate(v)
    If Err.Number <> 0 Then Err.Clear: DateJson = "null": Exit Function
    On Error GoTo 0
    DateJson = """" & Format(d, "yyyy-mm-dd") & """"
End Function

' Converte variante para String de forma segura:
' retorna "" para erros de célula (#N/A, #REF!, etc.), Empty e Null.
Private Function SafeStr(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then
        SafeStr = ""
    Else
        SafeStr = Trim$(CStr(v))
    End If
End Function

' Usa InStr para tolerar diferenças de codificação do 'ã' (NFC vs NFD) no Mac.
Private Function EhMovFinanceira(s As String) As Boolean
    Dim sl As String: sl = LCase$(s)
    EhMovFinanceira = (InStr(1, sl, "movimenta") > 0 And InStr(1, sl, "financeira") > 0)
End Function
