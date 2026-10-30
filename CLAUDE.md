# Zayda Construtora — Formulário de materiais + Painel de custos

Dois projetos do Google Apps Script, cada um num repositório e numa pasta própria.
Eles se comunicam **pelas planilhas**, não por código compartilhado.

| Pasta / repositório | Projeto Apps Script | Planilha a que está ligado | Arquivos |
|---|---|---|---|
| Formulario-Zayda-Construtora | ZAYDA CONSTRUTORA \| FORMULÁRIO (script 1TjrLZ…) | ZAYDA CONSTRUTORA \| CATÁLOGO E SOLICITAÇÕES (OFICIAL) | `Código.js`, `Index.html`, `appsscript.json` |
| Controle-Zayda-Construtora | Zayda Controle | ZAYDA CONSTRUTORA \| CONTROLE DE MATERIAIS POR OBRA (OFICIAL) | `Controle.js`, `Painel.html`, `appsscript.json` |

## Regras que não podem ser quebradas
- **Nunca renomear nem criar arquivos de código.** Os nomes são exatamente `Index.html`, `Código.js`, `Painel.html` e `Controle.js`. Nunca pode existir `INDEX.html` (o Git no Windows não distingue maiúsculas e isso já derrubou o formulário).
- **Conferir o `.clasp.json`**: o `scriptId` tem que ser o do projeto da tabela acima. Um ID errado envia o código para outro projeto (existem cópias de segurança com o mesmo nome).
- **Não mexer na planilha "Controle e Planejamento de Obras"**: o sistema só lê as abas OBRAS e ETAPAS dela.
- **Não mudar a ordem nem o significado das colunas existentes** das planilhas: há dados reais gravados. Campo novo = coluna nova no fim.
- **Não apagar dados** de PEDIDOS, ITENS_PEDIDO, ENTREGAS, AUDITORIA nem MATERIAIS_E_CUSTOS.
- Textos para o usuário em português do Brasil, datas dd/mm/aaaa, valores em R$, fuso America/Sao_Paulo.

## Sequência de trabalho (em cada pasta)
1. `git pull`
2. Editar e salvar.
3. `clasp push` → testar pelo link **/dev** (Implantar → Testar implantações).
4. `git add .` → `git commit -m "o que mudou"` → `git push`
5. Publicar: no Apps Script, Implantar → Gerenciar implantações → lápis → **Nova versão**. Isso é feito pelo usuário.

Mudança que envolve os dois projetos: fazer as alterações nos dois, e enviar/publicar **primeiro o formulário, depois o painel**.

## Como o dado anda (formulário → planilha → painel)
1. **Formulário** (`Index.html` + `Código.js`) grava na planilha do formulário:
   - `PEDIDOS` (sem cabeçalho): A PEDIDO_ID · B REVISAO · C CRIADO_EM · D ATUALIZADO_EM · E EMAIL_LOGIN · F ATENDENTE · G ENTREGADOR · H ENDERECO · I CASA_OBRA · J OBSERVACOES · K TOTAL_CONHECIDO · L ITENS_SEM_PRECO · M STATUS_PEDIDO · N ORIGEM · **O PONTO_REFERENCIA**. Próxima coluna livre: **P**.
   - `ITENS_PEDIDO` (11 colunas, sem cabeçalho), `ENTREGAS` (13 colunas, com cabeçalho: ... L RECEBIDO_POR, M ASSINATURA), `AUDITORIA`.
   - Lê `CATALOGO`, `CONFIG` e a aba `OBRAS`, que é **escrita pelo controle**: lista de obras para o campo Obra.
2. **Controle** (`Controle.js`, sincronização a cada minuto e a cada consulta do painel) lê a planilha do formulário e grava na planilha de controle:
   - `_IMPORT_PEDIDOS`: colunas A–N iguais a PEDIDOS, **O = OBRA_IDENTIFICADA** (calculada), **P = PONTO_REFERENCIA**. Próxima coluna livre: **Q**. Leitura em `lerOrigem_()`, escrita em `sincronizarFormulario()`.
   - `_IMPORT_ITENS`, `_IMPORT_CATALOGO`; entregas viram linhas `ENT-…` em `MATERIAIS_E_CUSTOS` (`lancarEntregas_`).
   - Obras e etapas: lidas da planilha de obras (`lerObras_`, `lerEtapas_`) → aba `OBRAS` e `_IMPORT_ETAPAS`.
3. **Painel** (`Painel.html`) recebe tudo de `getDadosPainel(codigo)` em `Controle.js` (`obras`, `itens`, `lanc`, `pedidos`, `nomesEtapas`).

**Campo novo do formulário que deve aparecer no painel** — mexer nos 5 pontos, nesta ordem:
1. `Index.html`: o campo, `payload()`, `openById`/`resetForm` e, se for impresso, `printRelation()`.
2. `Código.js`: `savePedido` (acrescentar no fim de `pedidoRow`, coluna P em diante) e `ultimos_` (ler a mesma coluna).
3. `Controle.js`: `lerOrigem_` (ler a coluna da PEDIDOS) e `sincronizarFormulario` (gravar na próxima coluna livre de `_IMPORT_PEDIDOS`, com cabeçalho).
4. `Controle.js`: `getDadosPainel` (ler a coluna e colocar em `pedidos[id]`).
5. `Painel.html`: mostrar onde for preciso (ao abrir o pedido em `vSol`, impressão em `folhaPedido`).

## Detalhes técnicos que costumam dar problema
- `Index.html` contém uma cópia das funções de identificação da obra (`semAcento_` … `resolverObra_`) que existem em `Controle.js`. **As duas cópias precisam continuar iguais.**
- `Painel.html` é um arquivo único com o logo embutido em base64 (`data:image/png;base64,…`): não reescrever essa linha. O painel não usa scriptlets (`<? ?>`); o `doGet` do controle troca o texto `"__URL_API__"` pelo endereço publicado.
- O painel é publicado "Executar como: Eu" + "Qualquer pessoa" e protegido por código de acesso (Script Property `CODIGO_PAINEL`). Dados chegam por `fetch` POST ao `doPost`, com `google.script.run` de reserva.
- O formulário é publicado "Executar como: usuário que acessa"; a lista de e-mails autorizados está em `CONFIG!EMAILS_PERMITIDOS`.
- Custo de cada entrega = preço do item no pedido (entra CONFIRMADO automaticamente).
- Nada de dados de exemplo em produção. Se precisar testar, usar atendente "TESTE".
