# Menu Lateral de Navegação para Excel — somente uso

Distribuição do suplemento COM/VSTO para instalar e usar no Excel. Inclui uma
planilha de teste editável, mas não inclui os fontes C# do suplemento.

O VBA da planilha de demonstração permanece aberto e editável, sem senha.
Para alterar o título, edite `tituloPainel` no início da macro
`MontarMenuDemoCompleta` no arquivo ou no módulo exportado.

## Conteúdo

- `deployment`: instalador e publicação VSTO assinada da versão 1.1.0.23.
- `demo/Preview_Funcional_MenuLateralNavegacao.xlsm`: planilha de teste.
- `demo/vba/DemoMenu.bas`: exportação do módulo VBA da demonstração.
- `.github/workflows/release.yml`: criação de uma GitHub Release ao enviar tag.

## Preparar o repositório no GitHub

1. Crie um repositório vazio no GitHub. Não selecione a opção para adicionar um
   README, `.gitignore` ou licença, pois estes arquivos já estão aqui.
2. No diretório deste projeto, inicialize o Git, adicione o repositório remoto,
   faça o primeiro commit e envie a branch padrão.
3. Crie e envie a tag `v1.1.0.23` para publicar a versão incluída:

   ```powershell
   git init
   git add .
   git commit -m "Initial release"
   git branch -M main
   git remote add origin https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git
   git push -u origin main
   git tag v1.1.0.23
   git push origin v1.1.0.23
   ```

4. No GitHub, acompanhe **Actions**. O workflow empacota o instalador e a
   planilha e cria uma Release com um ZIP para distribuição.

Para validar o empacotamento localmente:

```powershell
.\scripts\Package-Release.ps1 -Version v1.1.0.23 -OutputDirectory .\artifacts
```

O workflow usa a publicação assinada já presente em `deployment`; não recompila
nem reassina o add-in. Uma atualização VSTO exige uma publicação da nova versão
assinada com a mesma chave privada, que não deve ser enviada ao GitHub.

## Limites, assinatura e licença

O binário não vem acompanhado dos fontes C#. Isso evita incluir os arquivos
fonte no repositório, mas não impede tecnicamente engenharia reversa ou
alteração do binário. O certificado `.cer` é autoassinado; confira sua origem
e impressão digital antes de confiar nele. Os binários VSTO solicitam confiança
total; distribua somente a usuários que confiem na origem.

Nenhuma licença de software foi incluída. Defina uma licença apropriada antes
de autorizar redistribuição por terceiros.
