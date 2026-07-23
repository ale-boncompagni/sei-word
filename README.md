# SEI para Word

Extensão do Chrome para converter deliberações do SEI em formato Word (.docx).

## Descrição

Esta extensão permite converter deliberações do Sistema Eletrônico de Informações (SEI) diretamente para o formato Word (.docx), facilitando o compartilhamento e edição dos documentos.

## Instalação

1. Clone este repositório ou baixe os arquivos
2. Rode `npm install && npm run build` (gera `content.bundle.js` a partir de `src/`)
3. Abra o Chrome e acesse `chrome://extensions/`
4. Ative o "Modo do desenvolvedor" no canto superior direito
5. Clique em "Carregar sem compactação"
6. Selecione a pasta do projeto

## Uso

1. Acesse uma página de deliberação no SEI
2. Clique no botão de exportação para Word (ícone do Word) no menu do SEI
3. O arquivo será baixado automaticamente em formato .docx, nomeado como "Documento_SEI_YYYY-MM-DD.docx"

## Requisitos

- Google Chrome versão 88 ou superior
- Acesso ao SEI

## Desenvolvimento

O conversor (`src/converter/`) gera um `.docx` real (OOXML) com a biblioteca
`docx`, em vez do antigo truque de embutir HTML via MHTML/altChunk. Isso é o
que permite cabeçalho e rodapé nativos do Word, fontes/tamanhos corretos e
numeração de itens fiel ao SEI.

- `src/converter/htmlToDocx.js` — percorre o DOM do iframe do SEI e monta
  parágrafos, tabelas, cabeçalho e rodapé do Word.
- `src/converter/numberingRules.js` / `numberingEngine.js` — o SEI numera
  itens (`1.`, `1.1.`, `a)`, `I -`...) via `content: counter(...)` em CSS,
  que o Word não interpreta ao importar HTML. Este módulo reproduz a mesma
  lógica de `counter-increment`/`counter-reset` em JavaScript para gerar o
  texto real do número antes de montar o parágrafo.

Após editar `src/`, rode `npm run build` (ou `npm run watch`) antes de
recarregar a extensão em `chrome://extensions/`.

Para contribuir com o desenvolvimento:

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Faça commit das suas alterações (`git commit -am 'Adiciona nova feature'`)
4. Faça push para a branch (`git push origin feature/nova-feature`)
5. Crie um Pull Request

## Boas práticas e publicação

- O código segue as recomendações do Manifest V3
- Permissões mínimas: só `host_permissions` para os domínios do SEI atendidos (sem `activeTab`, sem `<all_urls>`)
- Não coleta, armazena ou transmite nenhum dado do usuário — toda a conversão roda localmente no navegador; a única rede acessada é a busca da imagem do cabeçalho, dentro dos próprios domínios já autorizados
- Sem código remoto: `content.bundle.js` é gerado localmente via `npm run build` e não carrega scripts de CDN nem usa `eval`
- Ícones (`icons/icon16.png`, `icon48.png`, `icon128.png`, `export-word-icon.png`) são originais, sem uso de marcas/logos de terceiros
- Não utiliza delays ou código duplicado
- O botão de exportação é inserido de forma dinâmica e segura
- O nome do arquivo é sempre seguro para qualquer sistema operacional

**Texto sugerido para a aba "Privacy practices" da Chrome Web Store:**
> Esta extensão não coleta, armazena ou transmite nenhum dado do usuário. Toda a conversão do documento do SEI para Word acontece localmente no navegador. O acesso aos domínios do SEI é necessário apenas para ler o conteúdo do documento aberto e montar o arquivo .docx.

## Licença

Este projeto está sob a licença MIT. 