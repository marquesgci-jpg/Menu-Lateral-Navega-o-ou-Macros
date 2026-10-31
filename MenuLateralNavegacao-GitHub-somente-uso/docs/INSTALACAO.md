# Instalação e ativação no Excel

Estas instruções são para o pacote Windows VSTO
`MenuLateralNavegacao-v1.1.0.23.zip`. Ele usa o mesmo certificado ClickOnce
das versões 1.1.0.1 a 1.1.0.22 e pode atualizá-las se mantiver o caminho da
instalação. Uma publicação ClickOnce antiga com outra identidade de assinatura
precisa ser removida separadamente.

## Preparar a confiança no certificado

O pacote inclui `MenuLateralNavegacao-Desenvolvimento.cer`, que contém apenas a
parte pública do certificado. O certificado é autoassinado: importar e confiar
nele é uma decisão explícita de segurança e não verifica a identidade civil ou
comercial do autor. Faça isso somente se recebeu o pacote de uma fonte em que
confia.

O certificado esperado tem o assunto
`CN=Menu Lateral de Navegacao - Desenvolvimento` e a impressão digital
`0E05C3A8B3F93AA42C8C26EBEFDDD15D13143441`. Para validar a origem, compare a
impressão digital com uma fonte independente da cópia do ZIP que está instalando.

1. Extraia todos os arquivos do ZIP para uma pasta local; não execute arquivos
   de dentro do ZIP.
2. Confira o assunto e a impressão digital do certificado antes de importá-lo.
   No Explorador de Arquivos, abra as propriedades do `.cer` e consulte a guia
   **Detalhes**; confirme a impressão digital.
3. Abra `certmgr.msc` e escolha **Ação > Todas as Tarefas > Importar**.
4. Importe o arquivo `.cer` para **Usuário Atual** e coloque-o em
   **Autoridades de Certificação Raiz Confiáveis**.
5. Repita a importação do mesmo `.cer` e coloque-o em
   **Editores Confiáveis**. Confira se o certificado corresponde ao pacote que
   recebeu antes de confiar nele.

Esses passos confiam neste certificado para o usuário atual do Windows. Não
instale ou compartilhe uma chave privada `.pfx`.

## Instalar

> A versão 1.1.0.23 inclui vistas configuráveis por ícones, pesquisa
> instantânea sem lupa, tema claro com destaque configurável, árvore sem
> moldura, ícones PNG de pastas fornecidos pelo usuário, botões de vista
> clássicos elevados (38×38 px) e título atualizável ao executar a macro
> novamente, preservando o conteúdo. No arquivo de demonstração, altere a
> variável `tituloPainel` no início de `MontarMenuDemoCompleta`; caracteres
> Unicode, incluindo acentos, são aceitos.
> Teste o menu
> em uma cópia da pasta de trabalho antes de usar em arquivos de produção.

1. Feche todas as janelas do Excel. A versão 1.1.0.23 pode atualizar a 1.1.0.22
   se usar o mesmo certificado e a publicação vier do mesmo caminho local.
   Se extrair em outra pasta, ou se houver uma instalação com identidade
   diferente, desinstale a versão existente primeiro em
   **Configurações > Aplicativos > Aplicativos instalados**.
2. Na pasta extraída, execute `setup.exe`. Se preferir, execute
   `Instalar_Suplemento.bat`, que abre o mesmo instalador.
3. Siga as etapas exibidas. Se o Windows solicitar elevação para instalar
   pré-requisitos, prossiga somente se tiver autorização e confiar na origem.
4. Se aparecer um aviso, confira se o publicador corresponde a
   **Menu Lateral de Navegacao - Desenvolvimento** e se a assinatura é válida.
   Não ignore avisos inesperados.
5. Abra o Excel novamente.

Para atualizações ClickOnce, extraia o novo ZIP sobre a mesma pasta usada na
instalação anterior, mantendo o caminho do arquivo `MenuLateralNavegacao.vsto`.
O ClickOnce não atualiza a partir de um caminho de origem diferente.

O instalador pode solicitar conexão com a internet para obter pré-requisitos,
como o runtime VSTO. O suplemento é uma extensão para o Excel desktop no
Windows; não é destinado ao Excel para a Web.

## Ativar ou conferir no Excel

1. No Excel, abra **Arquivo > Opções > Suplementos**.
2. No campo **Gerenciar**, selecione **Suplementos COM** e clique em **Ir...**.
3. Localize e marque **MenuLateralNavegacao** (ou nome semelhante) e confirme.
4. Volte à planilha e verifique se o menu lateral de navegação foi carregado.

Se o item não aparecer em **Suplementos COM**, confira também a lista **Itens
Desabilitados** na mesma tela de Suplementos. A interface e os nomes podem
variar ligeiramente conforme a versão do Office.

Para validar a instalação, use uma cópia da sua planilha de teste. A rotina
`GerarPainelACCONSULT` da planilha de exemplo apenas reconstrói os itens do
menu; não execute a opção de atualizar tabelas dinâmicas sem confirmar que isso
é desejado para os dados abertos.

O pacote 1.1.0.20 inclui a API para configurar vistas, tema claro, cor de destaque,
ícones e título opcional. Consulte a seção **Personalização do menu e aparência** em
[Reconstrução dos fontes C#](RECONSTRUCAO.md).

## Desinstalar

No Windows, abra **Configurações > Aplicativos > Aplicativos instalados**,
localize o suplemento e escolha **Desinstalar**. Se não houver uma entrada,
abra **Arquivo > Opções > Suplementos > Gerenciar: Suplementos COM > Ir...**,
desmarque o suplemento e consulte o administrador do computador para remover
a implantação VSTO.

## Solução de problemas

- **O instalador pede componentes adicionais:** permita que `setup.exe` obtenha
  os pré-requisitos oficiais ou peça ao administrador para instalá-los.
- **O suplemento não aparece:** feche e reabra o Excel, confira a caixa em
  **Suplementos COM** e verifique **Itens Desabilitados**.
- **Aviso de publicador ou certificado:** confirme a origem do ZIP e o nome do
  publicador. Verifique se o certificado público correto foi importado nos
  repositórios indicados; não ignore avisos inesperados.
- **Falha de instalação:** confirme que todos os arquivos do ZIP foram extraídos
  juntos e tente novamente com o Excel fechado. Se persistir, anote a mensagem
  completa do instalador para pedir suporte.

## Atualizações futuras

Uma atualização ClickOnce precisa ser assinada com a mesma chave privada desta
publicação. Faça uma cópia de segurança protegida por senha do PFX em local
seguro e fora do repositório; nunca compartilhe esse arquivo nem sua senha. Para
exportá-lo, abra `certmgr.msc` > **Pessoal > Certificados**, clique com o botão
direito no certificado indicado acima e escolha **Todas as Tarefas > Exportar**.
Selecione **Sim, exportar a chave privada**, o formato PFX e defina uma senha
forte. Se a opção de exportar a chave não estiver disponível, não há cópia de
segurança exportável. Se a chave for perdida, futuras versões precisarão ser
instaladas como uma nova identidade, após remover a anterior.
