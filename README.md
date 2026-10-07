# Gerador de Boletim Interno - COGER v3

A v3 acrescenta a estrutura hierárquica do Boletim Interno:

- **Título numerado:** `1.`, `2.`, `3.` ...
- **Subtítulo por letra:** `a.`, `b.`, `c.` ...
- títulos e subtítulos em **negrito**;
- a contagem dos subtítulos reinicia após cada novo título;
- as quatro partes do boletim aparecem no nível principal do índice;
- partes vazias recebem **SEM ALTERAÇÃO** e essa indicação também aparece no índice;
- cabeçalho das páginas seguintes com o texto à esquerda e `fl. X` alinhado à direita;
- índice Word real, com **pontilhado e número da página**;
- Word atualiza automaticamente o índice antes de salvar o DOCX e antes de exportar o PDF.

## Exemplo de estrutura

```text
3ª PARTE - ASSUNTOS GERAIS E ADMINISTRATIVOS

1. ORDEM DE MOVIMENTAÇÃO

a. ORDEM DE MOVIMENTAÇÃO Nº 096/2026 - COGER.
[texto]

b. ORDEM DE MOVIMENTAÇÃO Nº 097/2026 - COGER.
[texto]

2. NOTA PARA BOLETIM

a. NOTA PARA BOLETIM Nº 055/2026
[texto]
```

No índice, o Word gera automaticamente algo no formato:

```text
3ª PARTE - ASSUNTOS GERAIS E ADMINISTRATIVOS..............3
1. ORDEM DE MOVIMENTAÇÃO...................................3
   a. ORDEM DE MOVIMENTAÇÃO Nº 096/2026....................3
   b. ORDEM DE MOVIMENTAÇÃO Nº 097/2026....................4
2. NOTA PARA BOLETIM........................................5
   a. NOTA PARA BOLETIM Nº 055/2026.........................5
```

Os números são as páginas reais depois da paginação pelo Microsoft Word.

## Instalar

Execute `instalar.bat` ou:

```bat
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Executar

```bat
executar.bat
```

## Importante

O **Microsoft Word** precisa estar instalado para:

1. atualizar a numeração real do índice;
2. atualizar o campo `fl. X`;
3. gerar o PDF final.

O OCR continua usando o Tesseract e permanece com tela de revisão.


## v4 — Interface revisada

A v4 mantém a geração de DOCX/PDF e altera principalmente a experiência de uso:

- cabeçalho institucional mais limpo;
- campos gerais organizados em um único cartão;
- abas das quatro partes com visual mais claro;
- formulário e lista de publicações lado a lado;
- lista em formato de tabela com índice, tipo e título;
- duplo clique para editar uma publicação;
- botões separados para editar, mover e remover;
- confirmação antes de exclusão;
- contagem de caracteres;
- contador de publicações por parte e no boletim inteiro;
- barra de status;
- diálogos de OCR e assinaturas com o mesmo padrão visual;
- atalhos: `Ctrl+G` gera DOCX + PDF e `F6` verifica o OCR.

A lógica documental da v3 foi preservada.


## v5 — Interface moderna

A v5 muda apenas a apresentação e a experiência de uso, preservando a lógica documental da v4.

Principais mudanças:

- interface criada com CustomTkinter;
- barra lateral fixa com navegação pelas quatro partes;
- visual institucional mais limpo e atual;
- cards com cantos arredondados e melhor hierarquia visual;
- editor e lista de publicações lado a lado;
- botão principal de geração destacado;
- modo claro/escuro;
- resumo do BI no topo;
- tela de assinaturas modernizada;
- janelas de revisão OCR modernizadas;
- lista de publicações com edição por duplo clique;
- manutenção dos títulos, subtítulos, índice, paginação, OCR, DOCX e PDF da versão anterior.

Atalhos:
- Ctrl+G: gerar DOCX + PDF
- F6: verificar OCR

## v6 - Instalador com Tesseract embutido

Fluxo de compilacao no Windows:
1. Rode `instalar.bat`.
2. Tenha o Tesseract instalado no computador de desenvolvimento, incluindo Portugues.
3. Rode `preparar_tesseract.bat`.
4. Rode `gerar_exe_portatil.bat`.
5. Instale o Inno Setup 6.
6. Rode `gerar_instalador.bat`.

O resultado final sera:
`instalador\Instalador_Gerador_Boletim_COGER_v6.exe`

O PC de destino nao precisara instalar Python nem Tesseract. O Tesseract e a pasta `tessdata` serao instalados junto com o programa.

Observacao: a geracao de PDF e a atualizacao exata do indice ainda dependem do Microsoft Word, conforme a arquitetura atual do programa.


## v6.1 - Correção de permissão do Windows

Nesta versão, o aplicativo não tenta mais gravar o banco de dados dentro de
`C:\Program Files`.

O banco passa a ser salvo em:

`%LOCALAPPDATA%\Gerador Boletim COGER\database\boletim.db`

Os boletins gerados passam a ser salvos em:

`Documentos\Gerador Boletim COGER\Boletins`

Os arquivos internos do programa, como brasão e Tesseract empacotado,
continuam sendo lidos da pasta de recursos do executável.


## v6.2 - Instalação sem privilégios de administrador

Nesta versão, o instalador foi alterado para instalação por usuário.

O aplicativo será instalado em:

`%LOCALAPPDATA%\Programs\Gerador Boletim COGER`

Isso evita a necessidade de gravação em `C:\Program Files` e, por padrão,
não solicita elevação para administrador.

O instalador usa:

`PrivilegesRequired=lowest`

Os atalhos também são criados apenas para o usuário atual.

Para gerar:

1. `preparar_tesseract.bat`
2. `gerar_exe_portatil.bat`
3. `gerar_instalador.bat`

O instalador final será:

`instalador\Instalador_Gerador_Boletim_COGER_v6_2.exe`

Observação: políticas corporativas do Windows ainda podem bloquear instalação
ou execução de aplicativos não autorizados, mesmo sem solicitação de administrador.


## v6.3 - Word + LibreOffice automático

A aplicação agora detecta Microsoft Word e, se ele não estiver disponível, tenta usar LibreOffice automaticamente.

Ordem de preferência:
1. Microsoft Word
2. LibreOffice
3. Nenhum: gera DOCX, mas avisa que PDF/atualização automática não estão disponíveis.

O LibreOffice é procurado no PATH e também em `C:\Program Files\LibreOffice\program\soffice.exe` e `C:\Program Files (x86)\LibreOffice\program\soffice.exe`.

A instalação continua sem administrador e o Tesseract continua embutido.


## v6.4 - Padrão de data do Boletim

O período do boletim agora segue o padrão solicitado:

`01 out. à 07 out. 26`

Ou seja:
- a data inicial usa dia com dois dígitos + mês abreviado;
- a data final usa dia com dois dígitos + mês abreviado + ano com dois dígitos.

Exemplos:
- `01/10/2026` até `07/10/2026` -> `01 out. à 07 out. 26`
- `08/09/2026` até `14/09/2026` -> `08 set. à 14 set. 26`

O mesmo padrão é aplicado na capa e no cabeçalho das páginas seguintes.

Também foi removida uma chamada direta antiga ao Microsoft Word para preservar corretamente
o fallback automático Word -> LibreOffice introduzido na v6.3.


## v6.5 - Datas com calendário

Os campos **Data inicial** e **Data final** agora possuem botão de calendário.

Ao clicar no ícone de calendário:
- abre um calendário visual;
- o usuário escolhe o dia;
- a data é preenchida automaticamente no formato `DD/MM/AAAA`;
- o boletim continua sendo apresentado no padrão:
  `01 out. à 07 out. 26`.

Foi adicionada a dependência `tkcalendar` e os arquivos necessários foram
incluídos na configuração do PyInstaller.


## v6.6 - Documento somente em preto

A geração do DOCX/PDF foi ajustada para remover cores automáticas do Word.

Agora:
- títulos ficam pretos e em negrito;
- subtítulos ficam pretos e em negrito;
- texto normal fica preto;
- cabeçalho fica preto;
- índice fica preto;
- hyperlinks e entradas do índice não recebem azul.

O destaque visual do documento oficial é feito somente com negrito e,
quando aplicável, sublinhado.


## v6.7 - Correção RGBColor

Corrigido o erro:

`name 'RGBColor' is not defined`

O módulo `gerar_docx.py` agora importa explicitamente `RGBColor` de `docx.shared`,
mantendo a regra de documento oficial em texto preto.


## v6.8 - Salvamento automático e recuperação de rascunho

O boletim em edição agora é salvo automaticamente no banco SQLite local.

### Proteção contra perda de conteúdo

- salvamento automático a cada 3 segundos;
- último salvamento ao fechar normalmente;
- restauração automática na próxima abertura;
- salva publicações já adicionadas;
- salva também o texto que ainda está sendo digitado e que ainda não foi
  adicionado como publicação;
- preserva a parte atual;
- preserva título/subtítulo em edição;
- botão `Salvar rascunho agora`;
- botão `Novo boletim`, com confirmação antes de apagar o rascunho.

O rascunho é armazenado em:

`%LOCALAPPDATA%\Gerador Boletim COGER\database\boletim.db`

Portanto, uma falha do aplicativo não deve mais apagar o trabalho digitado.


## v6.9 - Histórico e edição de boletins antigos

A v6.9 adiciona um histórico permanente dos boletins.

### Novos recursos

- botão **Salvar boletim**;
- tela **Boletins salvos**;
- busca por número, local ou período;
- abrir um boletim antigo e continuar a edição;
- editar e salvar novamente o mesmo registro;
- duplicar um boletim para criar uma nova versão;
- excluir um boletim do histórico;
- gerar novamente DOCX + PDF diretamente pelo histórico;
- ao gerar um documento, o boletim também é salvo/atualizado automaticamente no histórico;
- o salvamento automático de rascunho da v6.8 continua funcionando.

Os boletins ficam armazenados no mesmo banco SQLite local:

`%LOCALAPPDATA%\Gerador Boletim COGER\database\boletim.db`

A exclusão no histórico não apaga DOCX ou PDF que já tenham sido gerados.
