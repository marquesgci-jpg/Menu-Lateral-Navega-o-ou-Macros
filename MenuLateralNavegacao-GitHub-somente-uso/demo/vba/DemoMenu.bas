Attribute VB_Name = "Módulo1"
' MENU LATERAL - MODELO EDITAVEL
' Todo o codigo de demonstracao fica neste unico modulo.
' Para adaptar ao seu projeto, comece pelo bloco CONFIGURACAO.
Option Explicit

'=========================================================================
' CONFIGURACAO: altere estes valores para personalizar o menu do seu projeto
' O titulo e definido pela variavel tituloPainel no inicio da macro principal.
'=========================================================================
Private Const ADDIN_ID As String = "MenuLateralNavegacao"
Private Const COR_DESTAQUE As String = "#0F766E"
Private Const NOME_PLANILHA_DEMO As String = "Planilha1"

' O exemplo troca somente a cor de destaque; o tema permanece claro.
Private corAlternativaAtiva As Boolean
Private tituloPainel As String

'=========================================================================
' INICIO RAPIDO
' Execute MontarMenuDemoCompleta para gerar o menu de exemplo.
' Personalize tituloPainel, as vistas, os icones e os itens abaixo.
'=========================================================================
Public Sub MontarMenuDemoCompleta()
    Dim menu As Object
    Dim ws As Worksheet
    Dim etapa As String
    Dim numeroErro As Long
    Dim descricaoErro As String

    ' Altere esta variavel para definir o titulo exibido no painel.
    tituloPainel = "Vida Vidinha Vidão"
    On Error GoTo Falha

    etapa = "Localizando o suplemento"
    Set menu = ObterPainel()
    Set ws = ThisWorkbook.Worksheets(NOME_PLANILHA_DEMO)

    etapa = "Definindo o titulo"
    menu.Title = tituloPainel

    etapa = "Limpando o menu anterior"
    menu.LimparMenu

    ' Cada vista aparece como um botao na barra superior do painel.
    ' Troque o texto e o nome do icone para personalizar os botoes.
    etapa = "Criando as vistas"
    menu.AdicionarVista "Visao geral", "Imagem 192"
    menu.AdicionarVista "Financeiro", "Imagem 228"
    menu.AdicionarVista "Fiscal", "Imagem 62"
    menu.AdicionarVista "Contabil", "Imagem 50"
    menu.AdicionarVista "Operacoes", "Imagem 160"
    menu.AdicionarVista "Relatorios", "Imagem 238"

    ' Em cada item, use: vista, pai, texto, macro, icone.
    ' Pai vazio cria uma pasta na raiz. Macro vazia cria apenas uma pasta.
    etapa = "Montando a vista inicial"
    menu.AdicionarItemNaVista "Visao geral", "", "COMECE POR AQUI", "", ""
    menu.AdicionarItemNaVista "Visao geral", "COMECE POR AQUI", _
        "Abrir central de demonstracao", "DemoAbrirDashboard", "Imagem 268"
    menu.AdicionarItemNaVista "Visao geral", "COMECE POR AQUI", _
        "Atualizar indicadores", "Macro_GerarDFC", "Imagem 244"
    menu.AdicionarItemNaVista "Visao geral", "", "NAVEGACAO E APARENCIA", "", ""
    menu.AdicionarItemNaVista "Visao geral", "NAVEGACAO E APARENCIA", _
        "Alterar cor de destaque", "DemoAlternarTema", "Imagem 258"
    menu.AdicionarItemNaVista "Visao geral", "NAVEGACAO E APARENCIA", _
        "Executar uma acao de exemplo", "Macro_AcaoDemo", "Imagem 252"

    etapa = "Montando a vista financeira"
    menu.AdicionarItemNaVista "Financeiro", "", "PLANEJAMENTO", "", ""
    menu.AdicionarItemNaVista "Financeiro", "PLANEJAMENTO", _
        "FLUXO DE CAIXA", "", "Imagem 238"
    menu.AdicionarItemNaVista "Financeiro", "FLUXO DE CAIXA", _
        "Consolidar demonstracao DFC", "Macro_GerarDFC", "Imagem 244"
    menu.AdicionarItemNaVista "Financeiro", "FLUXO DE CAIXA", _
        "Analisar variacao mensal", "Macro_GerarDFC", "Imagem 242"
    menu.AdicionarItemNaVista "Financeiro", "PLANEJAMENTO", _
        "CREDITO E RECEBIVEIS", "", "Imagem 252"
    menu.AdicionarItemNaVista "Financeiro", "CREDITO E RECEBIVEIS", _
        "Abrir simulador de credito", "Macro_Simulador", "Imagem 254"
    menu.AdicionarItemNaVista "Financeiro", "CREDITO E RECEBIVEIS", _
        "Consultar contas a receber", "Macro_AcaoDemo", "Imagem 262"

    etapa = "Montando a vista fiscal"
    menu.AdicionarItemNaVista "Fiscal", "", "OBRIGACOES TRIBUTARIAS", "", ""
    menu.AdicionarItemNaVista "Fiscal", "OBRIGACOES TRIBUTARIAS", _
        "Simples Nacional", "Macro_Simples", "Imagem 266"
    menu.AdicionarItemNaVista "Fiscal", "OBRIGACOES TRIBUTARIAS", _
        "Analise do Fator R", "Macro_FatorR", "Imagem 242"
    menu.AdicionarItemNaVista "Fiscal", "OBRIGACOES TRIBUTARIAS", _
        "Auditoria fiscal", "Macro_FatorR", "Imagem 264"
    menu.AdicionarItemNaVista "Fiscal", "", "DOCUMENTOS FISCAIS", "", ""
    menu.AdicionarItemNaVista "Fiscal", "DOCUMENTOS FISCAIS", _
        "Consultar documentos", "Macro_AcaoDemo", "Imagem 258"
    menu.AdicionarItemNaVista "Fiscal", "DOCUMENTOS FISCAIS", _
        "Exportar apuracao", "Macro_AcaoDemo", "Imagem 238"

    etapa = "Montando a vista contabil"
    menu.AdicionarItemNaVista "Contabil", "", "ESCRITURACAO", "", ""
    menu.AdicionarItemNaVista "Contabil", "ESCRITURACAO", _
        "Conciliacao bancaria", "Macro_GerarDFC", "Imagem 254"
    menu.AdicionarItemNaVista "Contabil", "ESCRITURACAO", _
        "Plano de contas", "Macro_AcaoDemo", "Imagem 258"
    menu.AdicionarItemNaVista "Contabil", "ESCRITURACAO", _
        "Folha de pagamento", "Macro_AcaoDemo", "Imagem 252"
    menu.AdicionarItemNaVista "Contabil", "", "CONTROLES PATRIMONIAIS", "", ""
    menu.AdicionarItemNaVista "Contabil", "CONTROLES PATRIMONIAIS", _
        "Ativos e depreciacao", "Macro_AcaoDemo", "Imagem 170"

    etapa = "Montando a vista de operacoes"
    menu.AdicionarItemNaVista "Operacoes", "", "ROTINAS INTERNAS", "", ""
    menu.AdicionarItemNaVista "Operacoes", "ROTINAS INTERNAS", _
        "Cadastro de clientes", "Macro_AcaoDemo", "Imagem 252"
    menu.AdicionarItemNaVista "Operacoes", "ROTINAS INTERNAS", _
        "Agendamento de tarefas", "Macro_AcaoDemo", "Imagem 240"
    menu.AdicionarItemNaVista "Operacoes", "ROTINAS INTERNAS", _
        "Backup em nuvem", "Macro_AcaoDemo", "Imagem 246"
    menu.AdicionarItemNaVista "Operacoes", "ROTINAS INTERNAS", _
        "Processos e protocolos", "Macro_AcaoDemo", "Imagem 270"
    menu.AdicionarItemNaVista "Operacoes", "", "ADMINISTRACAO", "", ""
    menu.AdicionarItemNaVista "Operacoes", "ADMINISTRACAO", _
        "Preferencias do projeto", "Macro_AcaoDemo", "Imagem 244"

    etapa = "Montando a vista de relatorios"
    menu.AdicionarItemNaVista "Relatorios", "", "PAINEIS GERENCIAIS", "", ""
    menu.AdicionarItemNaVista "Relatorios", "PAINEIS GERENCIAIS", _
        "Dashboard mensal", "Macro_GerarDFC", "Imagem 238"
    menu.AdicionarItemNaVista "Relatorios", "PAINEIS GERENCIAIS", _
        "Analise de desempenho", "Macro_FatorR", "Imagem 242"
    menu.AdicionarItemNaVista "Relatorios", "", "EXPORTACAO", "", ""
    menu.AdicionarItemNaVista "Relatorios", "EXPORTACAO", _
        "Relatorio fiscal", "Macro_Simples", "Imagem 266"
    menu.AdicionarItemNaVista "Relatorios", "EXPORTACAO", _
        "Central de documentos", "Macro_AcaoDemo", "Imagem 258"

    ' API classica: estes itens aparecem na arvore compartilhada.
    etapa = "Demonstrando a API classica"
    menu.AdicionarItem "", "ATALHOS DA API CLASSICA", ""
    menu.AdicionarItem "ATALHOS DA API CLASSICA", _
        "Executar atalho de exemplo", "DemoAcionarAtalho"

    etapa = "Aplicando a cor de destaque"
    menu.ConfigurarTema "Claro", COR_DESTAQUE

    etapa = "Atualizando a folha de orientacao"
    PrepararFolhaDemo ws
    corAlternativaAtiva = False
    ws.Activate
    ws.Range("A1").Select
    Exit Sub

Falha:
    numeroErro = Err.Number
    descricaoErro = Err.Description
    If Not ws Is Nothing Then
        ws.Range("Z1").Value = "Erro " & CStr(numeroErro) & ": " & _
            descricaoErro & " [" & etapa & "]"
    End If
    MsgBox "Nao foi possivel gerar o menu: " & descricaoErro & _
        " [" & etapa & "]", vbCritical, tituloPainel
End Sub

' Atalho com nome simples para executar a demonstracao pelo Excel.
Public Sub MontarMenuNovoDemo()
    MontarMenuDemoCompleta
End Sub

' Exemplo minimo da API, sem icones personalizados.
Public Sub TestarApiSemIcones()
    Dim menu As Object

    On Error GoTo Falha
    Set menu = ObterPainel()
    menu.Title = "Teste da API"
    menu.LimparMenu
    menu.AdicionarVista "Exemplo", ""
    menu.AdicionarItemNaVista "Exemplo", "", "MINHA AREA", "", ""
    menu.AdicionarItemNaVista "Exemplo", "MINHA AREA", _
        "Executar acao de exemplo", "Macro_AcaoDemo", ""
    menu.ConfigurarTema "Claro", COR_DESTAQUE
    Exit Sub

Falha:
    MsgBox "Erro no exemplo minimo: " & Err.Description, _
        vbCritical, "Teste da API"
End Sub

' Restaura o menu de demonstracao depois de testar a API minima.
Public Sub RestaurarMenuDemo()
    MontarMenuDemoCompleta
End Sub

'=========================================================================
' PERSONALIZE ESTAS MACROS: cada item do menu chama uma Sub publica daqui.
' Substitua o conteudo de exemplo pela rotina do seu proprio projeto.
'=========================================================================
Public Sub Macro_Simulador()
    MostrarAcaoDemo "Simulador de credito", _
        "Coloque aqui a logica do seu simulador e abra a tela correspondente."
End Sub

Public Sub Macro_GerarDFC()
    MostrarAcaoDemo "Demonstracao DFC", _
        "Coloque aqui a rotina que gera ou atualiza seu relatorio financeiro."
End Sub

Public Sub Macro_Simples()
    MostrarAcaoDemo "Simples Nacional", _
        "Coloque aqui o calculo ou a consulta fiscal do seu projeto."
End Sub

Public Sub Macro_FatorR()
    MostrarAcaoDemo "Analise do Fator R", _
        "Coloque aqui a rotina de calculo e exibicao do Fator R."
End Sub

Public Sub Macro_Refresh()
    ThisWorkbook.RefreshAll
    MsgBox "Atualizacao das consultas solicitada.", vbInformation, tituloPainel
End Sub

Public Sub Macro_AcaoDemo()
    MostrarAcaoDemo "Acao de exemplo", _
        "Substitua esta mensagem pela acao que o item do menu deve executar."
End Sub

'=========================================================================
' EXEMPLOS DE INTERACAO COM O MENU
'=========================================================================
Public Sub DemoAlternarTema()
    Dim menu As Object
    Dim cor As String

    On Error GoTo Falha
    Set menu = ObterPainel()

    If corAlternativaAtiva Then
        cor = COR_DESTAQUE
    Else
        cor = "#3B82F6"
    End If
    menu.ConfigurarTema "Claro", cor
    corAlternativaAtiva = Not corAlternativaAtiva
    Exit Sub

Falha:
    MsgBox "Nao foi possivel alterar a cor: " & Err.Description, _
        vbCritical, tituloPainel
End Sub

Public Sub DemoAbrirDashboard()
    With ThisWorkbook.Worksheets(NOME_PLANILHA_DEMO)
        .Activate
        .Range("A1").Select
    End With
End Sub

Public Sub DemoAcionarAtalho()
    MostrarAcaoDemo "Atalho da API classica", _
        "Este item foi criado com AdicionarItem, fora das vistas individuais."
End Sub

Public Sub LimparMenuDemo()
    Dim menu As Object

    On Error GoTo Falha
    Set menu = ObterPainel()
    menu.LimparMenu
    Exit Sub

Falha:
    MsgBox "Nao foi possivel limpar o menu: " & Err.Description, _
        vbCritical, tituloPainel
End Sub

'=========================================================================
' APOIO: normalmente nao e necessario alterar as rotinas desta secao.
'=========================================================================
Private Function ObterPainel() As Object
    Dim suplemento As Object

    Set suplemento = Application.COMAddIns(ADDIN_ID)
    If Not suplemento.Connect Then
        Err.Raise vbObjectError + 1100, "ObterPainel", _
            "O suplemento MenuLateralNavegacao nao esta conectado."
    End If

    Set ObterPainel = suplemento.Object
    If ObterPainel Is Nothing Then
        Err.Raise vbObjectError + 1101, "ObterPainel", _
            "O suplemento nao disponibilizou a interface de automacao."
    End If
End Function

Private Sub PrepararFolhaDemo(ByVal ws As Worksheet)
    ws.Range("A14").Value = _
        "Menu gerado. Altere tituloPainel no inicio da macro e execute novamente."
    ws.Range("A16").Value = "TITULO DO PAINEL"
    ws.Range("B16").Value = "Edite tituloPainel no inicio da macro."
    ws.Range("A17").Value = _
        "Altere a variavel tituloPainel em MontarMenuDemoCompleta."
    ws.Range("Z1").Value = "OK"
End Sub

Private Sub MostrarAcaoDemo(ByVal titulo As String, ByVal orientacao As String)
    MsgBox orientacao, vbInformation, titulo
End Sub

