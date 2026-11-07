# LibreMacros

Algumas macros para serem utilizadas no LibreOffice.

Hoje o repositório reúne macros em **Python** para o **LibreOffice Calc**, voltadas a quem trabalha com séries econômicas e gráficos:

| Arquivo | O que faz | Macros que aparecem no LibreOffice |
|---|---|---|
| [`sgs.py`](calc/python/sgs.py) | Baixa séries do **SGS (Banco Central)** e monta uma tabela com todas lado a lado | `AtualizarSGS` |
| [`ipeadata.py`](calc/python/ipeadata.py) | Baixa séries do **Ipeadata** (inclusive séries regionais, com uma coluna por território) | `AtualizarIpeadata` |
| [`formata_grafico.py`](calc/python/formata_grafico.py) | Aplica um **padrão de formatação aos gráficos** (título, subtítulo, fonte, legenda, grades, linha em y = 0…) | `GraficosFonte`, `GraficoTitulo`, `GraficoTituloEnglish`, `FormatarTodosGraficos`, `TituloETodosGraficos`, `RotulosPorMes` |
| [`diagnostico_rede.py`](calc/python/diagnostico_rede.py) | Mede onde o tempo vai numa requisição às APIs (DNS, conexão, servidor, download) | `DiagnosticoRede` |

## Início rápido

1. Copie os arquivos `.py` de `calc/python/` para a pasta de scripts Python do seu perfil do LibreOffice ([veja onde fica no seu sistema](#instalação)).
2. Abra o **Calc** e vá em **Ferramentas ▸ Macros ▸ Executar macro… ▸ Minhas macros ▸ `sgs` ▸ `AtualizarSGS`**. Na primeira vez, ela só cria as abas de parâmetros.
3. Preencha os códigos das séries na coluna A da aba `Parametros_SGS` e execute a macro de novo.
4. Para não repetir esse caminho de menus, [coloque a macro num botão ou atalho](#como-adicionar-à-interface).

---

## Sumário

1. [Requisitos](#requisitos)
2. [Instalação](#instalação) — [Windows](#windows) · [Linux](#linux) · [macOS](#macos)
3. [Como usar](#como-usar)
4. [Como adicionar à interface](#como-adicionar-à-interface) (menu, barra de ferramentas e atalhos)
5. [Configuração](#configuração)
6. [Solução de problemas](#solução-de-problemas)

---

## Requisitos

- **LibreOffice** com suporte a macros em Python. Desenvolvido e testado com o LibreOffice 24.2; em uso na 26.8.
- **Internet** para as macros do SGS e do Ipeadata.
- Nenhuma biblioteca extra: os scripts usam só o que já vem com o Python do LibreOffice (nada de `pandas`, `requests` etc.).

O suporte a Python **já vem** no LibreOffice para Windows e macOS. No Linux, depende do pacote da sua distribuição:

| Distribuição | Pacote a instalar (se ainda não estiver) |
|---|---|
| Debian / Ubuntu / Mint | `libreoffice-script-provider-python` |
| Fedora | `libreoffice-pyuno` |
| openSUSE | `libreoffice-pyuno` |
| Arch / Manjaro | já incluído nos pacotes `libreoffice-fresh` / `libreoffice-still` |
| Flatpak (Flathub) | já incluído |

Os nomes de pacote variam entre distribuições e versões; se o seu for diferente, procure por "pyuno" ou "script provider python".

---

## Instalação

Instalar é **copiar os arquivos `.py` para a pasta de scripts Python do seu perfil do LibreOffice**. Não há instalador.

### Passo comum: baixar os arquivos

- **Sem Git:** na página do repositório, clique em **Code ▸ Download ZIP** e extraia. Os arquivos estão em `calc/python/`.
- **Com Git:** `git clone https://github.com/pho-souza/LibreMacros.git`

### Como achar a pasta certa

A pasta de destino é `Scripts/python/` **dentro da pasta `user` do seu perfil do LibreOffice**. Se `Scripts` ou `python` não existirem, crie-as (nomes exatos, com o `S` maiúsculo em `Scripts`).

Para localizar a pasta `user`:

1. **Caminhos usuais.** São os listados em cada sistema, logo abaixo, e valem para a maioria das instalações.
2. **Pela janela de Opções (sem reiniciar).** Vá em **Ferramentas ▸ Opções ▸ LibreOffice ▸ Caminhos**, selecione uma linha como **Modelos** ou **Cópias de segurança (Backups)** e clique em **Editar…**. O caminho mostrado contém a pasta `user` do seu perfil, por exemplo `…/libreoffice/4/user/template`. A pasta onde `template` está é a `user`.
3. **Pelo modo de reparo.** **Ajuda ▸ Reiniciar em modo de reparo…** (em inglês, *Restart in Safe Mode*; em versões recentes o nome em inglês passou a ser *Troubleshoot Mode*). Na janela que aparece, procure o botão para mostrar o perfil do usuário (pode estar em **Avançado**). Atenção: isso reinicia o LibreOffice; para voltar ao normal, escolha a opção de reiniciar normalmente.

O caminho final fica assim: `<perfil>/user/Scripts/python/`.

Os caminhos usuais:

### Windows

1. Pressione **Win + R**, digite `%APPDATA%\LibreOffice\4\user` e tecle Enter.
2. Crie as pastas `Scripts` e, dentro dela, `python`, se ainda não existirem.
3. Copie para `python` os arquivos `.py` de `calc\python\`.

Caminho completo: `C:\Users\<você>\AppData\Roaming\LibreOffice\4\user\Scripts\python\`

> Instalou o LibreOffice pela **Microsoft Store**? O perfil fica num lugar diferente (dentro da pasta do aplicativo). Use o método **Ferramentas ▸ Opções ▸ Caminhos**, descrito acima.

### Linux

**Instalação nativa (apt, dnf, pacman…):**

```bash
mkdir -p ~/.config/libreoffice/4/user/Scripts/python
cp calc/python/*.py ~/.config/libreoffice/4/user/Scripts/python/
```

**Flatpak:**

```bash
mkdir -p ~/.var/app/org.libreoffice.LibreOffice/config/libreoffice/4/user/Scripts/python
cp calc/python/*.py ~/.var/app/org.libreoffice.LibreOffice/config/libreoffice/4/user/Scripts/python/
```

**Snap:** o perfil fica dentro de `~/snap/libreoffice/`; use o método **Ferramentas ▸ Opções ▸ Caminhos**, descrito acima.

**Dica para quem clona o repositório:** em vez de copiar, crie links simbólicos. Assim, um `git pull` já atualiza as macros:

```bash
ln -sf "$PWD"/calc/python/*.py ~/.config/libreoffice/4/user/Scripts/python/
```

> No Flatpak, o aplicativo roda num "sandbox". Links simbólicos para fora do sandbox podem não ser visíveis a ele; nesse caso, copie os arquivos.

### macOS

1. No Finder, tecle **Cmd + Shift + G** e cole:
   `~/Library/Application Support/LibreOffice/4/user`
2. Crie `Scripts` e, dentro dela, `python`, se ainda não existirem.
3. Copie para `python` os arquivos `.py` de `calc/python/`.

Ou, pelo Terminal:

```bash
mkdir -p ~/Library/Application\ Support/LibreOffice/4/user/Scripts/python
cp calc/python/*.py ~/Library/Application\ Support/LibreOffice/4/user/Scripts/python/
```

### Conferindo se deu certo

Abra o LibreOffice Calc e vá em **Ferramentas ▸ Macros ▸ Executar macro…**. Em **Minhas macros** devem aparecer os arquivos copiados (por exemplo `sgs`), e, ao selecionar um, suas macros na coluna da direita.

> O botão **Editar** não funciona para macros em Python: o LibreOffice só tem editor para Basic. Para alterar um script, edite o `.py` num editor de texto (VS Code, Kate, Notepad++…) e salve na mesma pasta. A extensão [APSO](https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python) facilita organizar e abrir scripts Python por dentro do LibreOffice (opcional).

---

## Como usar

Todas as macros abaixo são executadas em **Ferramentas ▸ Macros ▸ Executar macro… ▸ Minhas macros ▸ (arquivo) ▸ (macro) ▸ Executar**. Depois você pode [colocá-las em botões, menus e atalhos](#como-adicionar-à-interface).

### SGS (Banco Central) — `AtualizarSGS`

1. Abra (ou crie) a planilha em que quer os dados e execute **`AtualizarSGS`** uma primeira vez. Ela cria as abas `Parametros_SGS` e `Tabela_SGS` com os cabeçalhos.
2. Em `Parametros_SGS`, preencha a partir da linha 2:

   | Coluna | Conteúdo |
   |---|---|
   | **A** | Código da série no SGS (ex.: `189` IGP-M, `433` IPCA, `432` meta Selic) |
   | **B** | Descrição, opcional. Vira o título da coluna (sem ela, usa o código) |
   | **C** | Data inicial, opcional |
   | **D** | Data final, opcional |

3. Execute **`AtualizarSGS`** de novo. A tabela `Tabela_SGS` é preenchida com uma coluna de datas e uma coluna por série.

**Resultado e erros**

- A coluna **E** de `Parametros_SGS` mostra o status de cada linha, com cor: verde (ok), amarelo (aviso), vermelho (erro). Mostra também o nº de requisições, o tempo de rede e o tamanho baixado.
- Uma série com problema (código inexistente, API fora do ar…) **não derruba as outras** e **mantém a coluna que ela já tinha** na tabela. Se nenhuma série for baixada, a tabela não é alterada.
- A célula **F2** traz o resumo da execução: tempo de download, de gravação, de recálculo e total.
- Séries inexistentes **não recebem resposta** da API. Antes de baixar, a macro faz uma checagem rápida (pede só o último valor) e desiste após `TIMEOUT_EXISTENCIA_SEGUNDOS` (8 s por padrão), marcando o erro na coluna E.
- Sem datas, baixa a série inteira. Segundo o Portal de Dados Abertos do BCB, desde 26/03/2025 a API (formatos JSON e CSV) limita as consultas por período a **10 anos nas séries diárias**. Séries mensais e anuais (como o IGP-M) costumam vir inteiras numa só requisição; nas diárias, o pedido sem filtro é recusado e a macro refaz sozinha em janelas de `JANELA_ANOS` anos (3 por padrão), do período mais recente para o mais antigo, parando ao passar do início da série.
- Se você informar datas com intervalo maior que 10 anos, a macro também divide em janelas.
- Valores ausentes ficam em **branco**, não zero.

### Ipeadata — `AtualizarIpeadata`

Funciona como o SGS, com as abas `Parametros_Ipeadata` e `Tabela_Ipeadata`:

| Coluna | Conteúdo |
|---|---|
| **A** | Código da série (`SERCODIGO`), ex.: `PRECOS12_IPCA12`. Aceita minúsculas e espaços |
| **B** | Descrição (opcional) |
| **C** / **D** | Data inicial / final (opcionais) |
| **E** / **F** | Status por linha / log geral (preenchidos pela macro) |

- A API devolve a série inteira; o filtro de datas é aplicado depois do download.
- Séries com **vários territórios** (estados, regiões…) ganham uma coluna por território: `descrição [código do território]` (limite de 100 territórios por série).

### Download em segundo plano

Por padrão (`EXECUCAO_ASSINCRONA = True`), as duas macros baixam em segundo plano: o LibreOffice **não trava**, você pode continuar usando a planilha, e a gravação acontece sozinha no fim, com um aviso na barra de status. Durante o download, a célula F2 mostra "Baixando em segundo plano…".

- Chamar a macro de novo durante a execução é ignorado.
- Evite editar `Parametros_*` durante o download (as mudanças só valem na próxima execução) e evite editar a tabela de destino, que será reescrita no fim.
- Prefere que a macro espere terminar? Coloque `EXECUCAO_ASSINCRONA = False` no topo do arquivo.

### Formatação de gráficos — `formata_grafico.py`

Aplica um padrão visual ao gráfico: Calibri 10, título em caixa-alta e negrito, subtítulo, fundo, área de plotagem, grades, legenda embaixo na largura total, caixa "Fonte" no rodapé, marcas externas nos eixos, e uma **linha preta em y = 0** quando o gráfico tem valores positivos e negativos.

**Qual gráfico é formatado**

1. Se há gráfico selecionado (um clique nele), usa o selecionado.
2. Senão, se a aba tem **exatamente um** gráfico, usa esse.
3. Senão, pede para você selecionar um gráfico. Nada é alterado sem uma dessas condições.

**Macros**

| Macro | O que faz |
|---|---|
| `GraficosFonte` | Só formata, sem ler células. |
| `GraficoTitulo` | Lê **A3** (título), **A4** (subtítulo) e **A5** (fonte), liga os nomes das séries à linha 7 e depois formata. |
| `GraficoTituloEnglish` | Igual, em inglês: título em A4, séries na linha 6, procura "Source". |
| `FormatarTodosGraficos` | `GraficosFonte` em todos os gráficos do arquivo. |
| `TituloETodosGraficos` | `GraficoTitulo` em todos os gráficos do arquivo (cada aba lê o próprio A3/A4/A5). |
| `RotulosPorMes` | Rótulos de dados só nos pontos de um mês (pergunta o número do mês, de 1 a 12). |

**Regras para A3, A4 e A5** (`GraficoTitulo`)

- **A3:** título. Se contiver número ou data, **não** vira título: aparece como aviso.
- **A4:** se tiver texto, vira o subtítulo, colado ao título. Em branco, é ignorada (um subtítulo que já exista no gráfico não é removido).
- **A5:** se estiver preenchida, vira o texto da caixa "Fonte". Se não começar com "Fonte", a macro acrescenta `Fonte: ` na frente. Em branco, é ignorada.
- O layout é dinâmico: sem subtítulo, a área de plotagem sobe; ela sempre termina acima da legenda.

A mensagem ao final lista o título, subtítulo e fonte que foram lidos e quaisquer avisos, por aba.

### Diagnóstico de rede — `DiagnosticoRede`

Se as macros do SGS ou do Ipeadata estiverem lentas, execute **`DiagnosticoRede`**. Ela escreve um relatório na aba `Diagnostico_rede` com o tempo de DNS, de conexão em cada endereço (IPv4/IPv6), de espera do servidor, de download e de leitura, e aponta a etapa mais lenta. Útil, por exemplo, quando um dos endereços do servidor está fora do ar e só parte das conexões trava.

---

## Como adicionar à interface

Você pode chamar as macros de três formas, todas em **Ferramentas ▸ Personalizar…**. Escolha sempre o escopo **LibreOffice Calc** (em versões antigas, o campo se chama **Salvar em**) e não "este documento", para a personalização valer em qualquer arquivo.

> Os nomes exatos dos botões e campos dessa janela variam um pouco entre as versões do LibreOffice; os passos abaixo descrevem o caminho geral.

### Botão numa barra de ferramentas

1. **Ferramentas ▸ Personalizar… ▸ aba Barras de ferramentas**.
2. Ao lado da lista **Barra de ferramentas de destino**, abra o menu do botão de engrenagem (ou o botão **Adicionar**), escolha **Adicionar…**, dê um nome (por exemplo, `Macros`) e confirme o escopo **LibreOffice Calc**.
3. Com a barra nova selecionada, clique no botão **Adicionar…** (ou **+**) do conteúdo da barra, à direita.
4. Em **Categoria**, escolha **Macros do LibreOffice ▸ Minhas macros ▸ (arquivo)** e, em **Comando**, a macro desejada (por exemplo, `AtualizarSGS`). Clique em **Adicionar**.
5. Selecione o botão na lista e use **Modificar ▸ Renomear** e **Alterar ícone** para dar um nome curto e um ícone.
6. **OK**. Se a barra não aparecer, ative-a em **Exibir ▸ Barras de ferramentas**.

> Usa a interface **Barra de guias** (abas Início, Inserir… com faixa de opções)? As abas dessa interface **não são personalizáveis** pelo menu. A barra de ferramentas criada acima continua funcionando: ela aparece junto das demais barras, ativada em **Exibir ▸ Barras de ferramentas**. Também dá para voltar à interface clássica em **Exibir ▸ Interface do usuário ▸ Padrão**.

### Item de menu

1. **Ferramentas ▸ Personalizar… ▸ aba Menus**.
2. Em **Destino**, escolha **LibreOffice Calc** e selecione o menu em que quer o item (por exemplo, **Ferramentas**).
3. Clique em **Adicionar…** (ou no botão **+**), escolha **Macros do LibreOffice ▸ Minhas macros ▸ (arquivo)** e a macro, e **Adicionar**.
4. Se quiser um menu próprio (por exemplo, "RAF"), use o botão **Novo menu** da mesma janela e coloque as macros dentro dele.

### Atalho de teclado

1. **Ferramentas ▸ Personalizar… ▸ aba Teclado**.
2. Escolha **LibreOffice** (vale para todos os módulos) ou **Calc**.
3. Na lista de atalhos, clique na combinação que quer usar. A janela avisa se ela já está em uso.
4. Em **Categoria**, escolha **Macros do LibreOffice ▸ Minhas macros ▸ (arquivo)** e, em **Função**, a macro.
5. Clique em **Modificar** e depois em **OK**.

Sugestão de atalhos para as macros de gráfico (altere à vontade):

| Atalho | Macro |
|---|---|
| `Ctrl + Alt + I` | `GraficosFonte` |
| `Ctrl + Alt + J` | `GraficoTitulo` |

> Combinações com `Ctrl + Alt` costumam ser reservadas pelo ambiente gráfico no Linux, e teclas como `Ç` podem não ser reconhecidas em alguns layouts de teclado. Teste cada atalho depois de definir.

### Exportar e levar a configuração para outro computador

Sua barra, menus e atalhos ficam guardados no perfil do LibreOffice. Para repetir a configuração em outra máquina, instale os scripts (veja [Instalação](#instalação)) e refaça os passos acima. Para quem prefere empacotar tudo de uma vez, o caminho é criar uma extensão `.oxt` (não incluída aqui).

---

## Configuração

Cada script tem, **no topo do arquivo**, um bloco de constantes que você pode editar num editor de texto. As principais:

**`sgs.py` e `ipeadata.py`**

| Constante | Padrão | Para que serve |
|---|---|---|
| `EXECUCAO_ASSINCRONA` | `True` | Baixa em segundo plano, sem travar a planilha |
| `AVISO_FINAL` | `"BARRA"` | Aviso ao terminar em segundo plano: `"BARRA"`, `"CAIXA"` ou `None` |
| `TIMEOUT_SEGUNDOS` | `30` (SGS) / `20` (Ipeadata) | Espera máxima por resposta **de dados** (uma janela de série diária pode demorar). No Ipeadata, é também o tempo gasto por um código inexistente |
| `TIMEOUT_EXISTENCIA_SEGUNDOS` | `8` (só SGS) | Espera máxima pela checagem "a série existe?". É o tempo gasto por uma série inexistente |
| `VERIFICAR_EXISTENCIA` | `True` (só SGS) | Faz a checagem rápida antes de baixar (uma requisição pequena a mais por série) |
| `JANELA_ANOS` | `3` (só SGS) | Tamanho das janelas quando a API recusa o pedido sem filtro (séries diárias). O máximo permitido pelo BCB é 10 |
| `MAX_THREADS` | `4` | Downloads simultâneos |
| `CONEXAO_RAPIDA_SEGUNDOS` | `2.0` | Espera por cada endereço do servidor ao conectar (evita travar num endereço fora do ar) |
| `PREFERIR_IPV4` | `True` | Tenta o IPv4 antes do IPv6 nesses servidores |
| `MAX_LINHAS_PARAMETROS` | `200` | Quantas linhas de `Parametros_*` são lidas |

**`formata_grafico.py`**

| Constante | Padrão | Para que serve |
|---|---|---|
| `FONTE`, `TAM_TEXTO`, `TAM_TITULO` | Calibri, 10, 14 | Tipografia |
| `COR_FUNDO` / `PAREDE_COR` | `#F5F5F5` / `#EEEEEE` | Fundo do gráfico / da área de plotagem |
| `COR_GRADE`, `ESPESSURA_GRADE` | `#757575`, 0,01 cm | Grade principal |
| `POSICAO_LEGENDA`, `LEGENDA_LARGURA_TOTAL` | `"BOTTOM"`, `True` | Legenda |
| `LINHA_ZERO`, `LINHA_ZERO_MODO` | `True`, `"EIXO"` | Linha preta em y = 0. O modo `"EIXO"` sobrevive ao salvar em `.xlsx`; `"SECUNDARIO"` não |
| `CELULA_TITULO`, `CELULA_SUBTITULO`, `CELULA_FONTE` | `A3`, `A4`, `A5` | Células lidas por `GraficoTitulo` |
| `PREFIXAR_FONTE` | `True` | Acrescenta `Fonte: ` ao texto de A5 se faltar |

Depois de editar um arquivo, salve-o na pasta de scripts. Se a mudança não surtir efeito, feche e abra o LibreOffice.

---

## Solução de problemas

**As macros não aparecem em "Minhas macros".**
Confira se os arquivos estão em `…/user/Scripts/python/` do perfil certo (veja [Como achar a pasta certa](#como-achar-a-pasta-certa-vale-para-qualquer-sistema)). No Linux, confirme que o pacote de suporte a Python está instalado ([Requisitos](#requisitos)). Feche e abra o LibreOffice.

**O botão "Editar" não abre nada.**
É esperado para Python. Edite o `.py` num editor de texto.

**A macro do SGS ou do Ipeadata demora.**
Execute `DiagnosticoRede`. Depois de rodar uma atualização, olhe também a célula **F2** (onde o tempo foi: download, gravação, recálculo) e a coluna **E** (nº de requisições e novas tentativas). Uma série **inexistente** custa um timeout, porque a API simplesmente não responde; a linha mostra `ERRO: sem resposta do servidor`.

**`ERRO: sem resposta do servidor em 8 s (a série existe? …)`.**
No SGS, a checagem de existência não recebeu resposta: a série provavelmente não existe (ou o servidor está fora do ar). Confira o código.

**`ERRO: sem resposta do servidor em 30 s (consulta pesada ou servidor lento …)`.**
A série existe, mas a consulta de dados demorou mais que `TIMEOUT_SEGUNDOS`. Em séries diárias longas, aumente `TIMEOUT_SEGUNDOS` ou diminua `JANELA_ANOS`, para que cada janela seja menor.

**"A3 contém um número ou data… ignorei essa célula" (macros de gráfico).**
O título é lido de A3 e não pode ser número ou data. Escreva o título como texto. O mesmo vale para A4 (subtítulo).

**"Há N gráficos nesta aba. Clique uma vez no gráfico desejado…"**
A aba tem mais de um gráfico e nenhum está selecionado. Clique uma vez no gráfico e rode a macro de novo.

**"Não consegui acessar os dados internos do gráfico."**
O LibreOffice às vezes entrega o gráfico sem dados logo ao abrir um arquivo. Rode a macro de novo; se persistir, salve, feche e reabra o arquivo.

**Quero que o subtítulo (A4) ou a fonte (A5) sejam ignorados.**
Deixe a célula vazia (selecione e tecle **Delete**). Uma célula só com espaços também é tratada como vazia. Lembre que um subtítulo ou fonte que já exista no gráfico não é removido quando a célula está em branco.

**No Flatpak, a macro não alcança uma pasta do meu computador.**
O Flatpak é isolado. Conceda acesso à pasta com `flatpak override --user --filesystem=/caminho/da/pasta org.libreoffice.LibreOffice`.

---

## Contribuindo

Sugestões e correções são bem-vindas: abra uma *issue* descrevendo o problema (versão do LibreOffice, sistema operacional, e o texto da célula **F2** ou da coluna **E** quando for o caso) ou envie um *pull request*.
