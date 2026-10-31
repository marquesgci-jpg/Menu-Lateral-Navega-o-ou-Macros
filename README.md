# Menu Lateral de Navegação para Excel

Suplemento VSTO para Microsoft Excel que disponibiliza um menu lateral de navegação.
Este repositório organiza a distribuição atualmente disponível do suplemento.

> **Estado do código-fonte:** o projeto-fonte original não foi encontrado. Os
> arquivos em `src/` são uma reconstrução por descompilação da DLL publicada.
> O pacote com fontes inclui esses arquivos para consulta e modificação, mas
> eles não são os fontes originais do projeto. Consulte
> [Reconstrução dos fontes C#](docs/RECONSTRUCAO.md).

## Baixar e instalar

Baixe o pacote da versão mais recente na página
[Releases](../../releases/latest). Extraia o ZIP e siga as instruções em
[Instalação e ativação no Excel](docs/INSTALACAO.md).

Há duas distribuições com a planilha de teste:

- `dist/MenuLateralNavegacao-v1.1.0.23-com-fontes.zip`: inclui o instalador,
  a planilha e os fontes C# reconstruídos do suplemento.
- `dist/MenuLateralNavegacao-v1.1.0.23-somente-uso.zip`: inclui o instalador e
  a planilha, mas não inclui os fontes do suplemento.

Em ambas, o projeto VBA da planilha permanece aberto e editável. A diferença
entre os pacotes é somente a inclusão dos fontes do suplemento. O pacote
somente uso não impede tecnicamente engenharia reversa dos binários.

Esta é uma reconstrução para validação. Os manifestos do pacote 1.1.0.23 foram
assinados com a identidade existente. A árvore usa os PNGs de pasta aberta e
fechada enviados pelo usuário, em 32×32 px, sem moldura. Os botões das vistas
usam o estilo padrão elevado do Windows, com ícones centralizados de 24×24 px
em botões de 38×38 px. O título do painel pode ser alterado ao executar
novamente a macro, sem perder as vistas e os itens. Na planilha de demonstração, altere a variável `tituloPainel` no início da
macro `MontarMenuDemoCompleta` e execute-a novamente. A variável aceita
caracteres Unicode, incluindo acentos. A busca não exibe mais a lupa e o botão
de recolhimento tem uma área de clique maior.
Valide as vistas, os ícones, a cor de destaque e o título opcional em uma cópia
da pasta de trabalho antes de distribuir amplamente.

## Personalizar o painel

As pastas de trabalho podem definir vistas, itens, macros e ícones em suas
próprias Subs VBA. `Title` define o cabeçalho nativo do painel quando ele é
criado; o mesmo valor pode ser reaplicado ao reconstruir o menu. `ConfigurarTema`
usa o tema claro e aceita uma cor de destaque hexadecimal. A árvore usa o
estilo clássico do Windows, com pastas amarelas e ícones de página. Imagens
opcionais podem ser incorporadas na planilha `RecursosMenu` e referenciadas
pelo nome do objeto. O rodapé com o nome do autor e o link Patreon é fixo no
suplemento. Consulte [Reconstrução dos fontes C#](docs/RECONSTRUCAO.md) para a
API e um exemplo VBA.

## Requisitos

- Windows com a versão desktop do Microsoft Excel.
- .NET Framework 4.7.2 e o runtime do Visual Studio Tools for Office (VSTO).
  O instalador de bootstrap (`setup.exe`) pode verificar/instalar pré-requisitos.
- Conexão com a internet se o instalador precisar obter pré-requisitos.

## Estrutura

```text
.
├── Application Files/          # Publicação ClickOnce do suplemento
├── docs/
│   ├── INSTALACAO.md           # Instalação, ativação e solução de problemas
│   ├── RECONSTRUCAO.md         # Proveniência e limitações dos fontes recuperados
│   └── PROMPT_THUMBNAIL.md     # Prompt para criar thumbnail de divulgação
├── dist/
│   ├── MenuLateralNavegacao-v1.1.0.23-com-fontes.zip
│   └── MenuLateralNavegacao-v1.1.0.23-somente-uso.zip
├── src/
│   └── MenuLateralNavegacao/   # Fontes C# recuperados da DLL publicada
├── Instalar_Suplemento.bat     # Abre o instalador incluído
├── MenuLateralNavegacao.vsto   # Manifesto de implantação VSTO
├── readme.txt                  # Atalho para a documentação
└── setup.exe                   # Bootstrapper do instalador
```

As planilhas de exemplo, logs e a cópia antiga `MenuLateralNavegacao_Arquivos`
foram excluídos do pacote de instalação e não são necessárias para compilar o
suplemento.

## Assinatura e confiança

O pacote inicial usa um certificado autoassinado para
`Menu Lateral de Navegacao - Desenvolvimento`. Isso não comprova a identidade
do desenvolvedor para terceiros: cada usuário deve conferir a origem do
certificado `.cer` e decidir explicitamente se confia nele. As instruções de
instalação explicam como confiar no certificado para o usuário atual do Windows.
O certificado esperado tem a impressão digital
`0E05C3A8B3F93AA42C8C26EBEFDDD15D13143441`; compare-a com uma fonte independente
antes de confiar no pacote. Não distribua o arquivo PFX nem a chave privada.

O certificado e sua chave privada precisam ser preservados para assinar futuras
versões com a mesma identidade. Perder a chave exigirá uma nova identidade e
uma nova instalação, não uma atualização ClickOnce.

Os binários declaram permissões de confiança total. Distribua-os somente a
usuários que confiem na origem e depois de validar a assinatura e o conteúdo
do pacote.

## Preparar uma versão

1. Teste a DLL reconstruída no Excel e valide a publicação VSTO.
2. Crie o ZIP com `setup.exe`, `MenuLateralNavegacao.vsto` e a pasta
   `Application Files` da mesma versão.
3. Nomeie o arquivo `MenuLateralNavegacao-v<versão>.zip`.
4. Publique o ZIP como asset de uma nova GitHub Release e atualize o changelog.

Os fontes atuais são uma recuperação, não uma cópia do projeto-fonte original.
Antes de aceitar contribuições de desenvolvimento ou substituir o instalador,
valide a compilação, a integração VSTO e uma nova publicação assinada.
