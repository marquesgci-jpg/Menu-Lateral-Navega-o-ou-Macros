# Menu Lateral de Navegação para Excel

Suplemento COM/VSTO que adiciona ao Excel um painel lateral para organizar e
executar as macros de uma pasta de trabalho. Em vez de procurar comandos em
abas, botões ou na lista de macros, o usuário pode navegar por vistas, grupos
hierárquicos e uma caixa de pesquisa.

## Veja o suplemento

![(docs/images/menu-lateral-final.png)](https://github.com/marquesgci-jpg/Menu-Lateral-Navega-o-ou-Macros/blob/main/MenuLateralNavegacao-GitHub-somente-uso/docs/images/menu-lateral-final.png)

## O que você pode fazer

- Criar vistas independentes, acessadas por botões com ícones.
- Organizar comandos em uma árvore com pastas e subpastas.
- Pesquisar os itens da vista ativa.
- Associar itens a macros públicas do próprio arquivo Excel.
- Escolher o título do painel e uma cor de destaque em tema claro.
- Manter a lógica do projeto em VBA: o suplemento fornece a interface e chama
  as macros que você associar aos itens.

## Instalação e primeiro uso

Requisitos: Windows, Excel desktop e o runtime VSTO/.NET Framework. O Excel
para a Web não carrega suplementos COM/VSTO.

1. Baixe o pacote ZIP na seção **Releases** e extraia todos os arquivos para
   uma pasta local.
2. Confira a origem do pacote e do certificado antes de confiar nele. O
   certificado incluído é autoassinado; veja [Instalação](docs/INSTALACAO.md)
   para conferir a impressão digital e os passos completos.
3. Execute `deployment/setup.exe` e conclua a instalação.
4. Abra `demo/Preview_Funcional_MenuLateralNavegacao.xlsm` no Excel e confirme
   que **MenuLateralNavegacao** está conectado em **Arquivo > Opções >
   Suplementos > Gerenciar: Suplementos COM > Ir...**.
5. Execute `MontarMenuDemoCompleta` para montar o menu de exemplo.

A planilha contém ações demonstrativas. Antes de usar o suplemento em um
projeto real, substitua as macros de exemplo pelas rotinas do seu arquivo.
As macros VBA da planilha permanecem abertas e editáveis, sem senha.

## Personalizar o menu

No VBA da planilha, abra `Módulo1` e altere a variável `tituloPainel` no início
de `MontarMenuDemoCompleta`. Execute a macro novamente para atualizar o painel.
O valor pode conter acentos e outros caracteres Unicode.

Para adaptar a navegação:

1. Troque as chamadas `AdicionarVista` para definir os botões e seus ícones.
2. Use `AdicionarItemNaVista` para criar pastas, subpastas e comandos. Um pai
   vazio cria um item na raiz; informe o nome da pasta em `pai` para aninhar.
3. Associe comandos a uma `Public Sub` sem parâmetros do mesmo arquivo VBA.
   Essa é a rotina chamada quando o usuário seleciona o item.
4. Ajuste `ConfigurarTema "Claro", "#RRGGBB"` para escolher a cor de destaque.

Exemplo mínimo:

```vb
Public Sub MontarMeuMenu()
    Dim menu As Object
    Set menu = Application.COMAddIns("MenuLateralNavegacao").Object

    menu.Title = "Meu projeto"
    menu.LimparMenu
    menu.ConfigurarTema "Claro", "#0F766E"
    menu.AdicionarVista "Geral", "Imagem 268"
    menu.AdicionarItemNaVista "Geral", "", "NAVEGACAO", "", ""
    menu.AdicionarItemNaVista "Geral", "NAVEGACAO", _
        "Abrir painel", "AbrirPainel", "Imagem 268"
End Sub

Public Sub AbrirPainel()
    MsgBox "Substitua esta mensagem pela rotina do seu projeto."
End Sub
```

Veja a planilha de demonstração e [a descrição da API recuperada](docs/RECONSTRUCAO-FONTES.md)
para exemplos adicionais.

## Esta edição: com fontes

Inclui o projeto C# em `src/MenuLateralNavegacao` e os recursos gráficos, além
da publicação instalável, da planilha de teste e do módulo VBA exportado.
Os fontes C# são uma reconstrução por descompilação da DLL publicada; o
projeto original não foi localizado e pode diferir desta reconstrução.

O workflow gera um pacote de Release ao publicar uma tag correspondente à
versão da implantação assinada que está em `deployment`. Ele empacota essa
publicação existente, mas não recompila nem reassina o add-in. Atualizações
VSTO precisam da mesma chave privada de assinatura, que não está neste projeto
e nunca deve ser enviada ao GitHub.

## Confiança e distribuição

Os binários VSTO solicitam confiança total. Verifique a origem do pacote e do
certificado antes de instalar. Não compartilhe a chave privada `.pfx`.
Nenhuma licença de software foi incluída; defina uma licença antes de
redistribuir ou aceitar contribuições.
