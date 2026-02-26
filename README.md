# Gerador de XML — Planilhas de Eventos

Ferramenta desktop em Python para leitura de planilhas Excel (`.xlsx`) com layout de eventos posicionais, edição interativa de campos, validação de posicionamento e geração automática de arquivos XML, SQL e cópia da planilha para integração com sistemas corporativos.

---

## Sumário

1. [Visão Geral](#visão-geral)
2. [Requisitos](#requisitos)
3. [Como Executar](#como-executar)
4. [Interface](#interface)
5. [Fluxo de Uso](#fluxo-de-uso)
6. [Estrutura da Planilha de Entrada](#estrutura-da-planilha-de-entrada)
7. [Arquivos Gerados](#arquivos-gerados)
8. [Regras de Negócio](#regras-de-negócio)
9. [Validação](#validação)
10. [Atalhos de Teclado](#atalhos-de-teclado)

---

## Visão Geral

O sistema lê uma planilha principal com a definição dos campos de um evento posicional (layout de registro), permite ao usuário editar, copiar campos de outra planilha (origem) e gerar automaticamente os seguintes artefatos:

| Arquivo gerado | Conteúdo |
| --- | --- |
| `{ID}_Layout_entrada.xml` | Layout posicional completo (todos os campos ativos) |
| `{ID}_Layout_persistencia.xml` | Campos marcados com `Persistencia=S` |
| `{ID}_Layout_mapa_atributo.xml` | Campos marcados com `MapaAtributo=S` (namespace ns2) |
| `{ID}_Layout_enriquecimento.xml` | Dados de enriquecimento (DadoExterno) com CDATA |
| `ComandoSQL.sql` | Scripts INSERT para tabela `COLUMN_CONFIGURATION` |
| `evento_event_{ID}.xlsx` | Cópia estruturada da planilha com todas as alterações |

> `{ID}` = campo `IdentificadorEvento` da aba "Identificação Evento" da planilha.

---

## Requisitos

- **Python 3.8+**
- **openpyxl >= 3.0.10**
- Tkinter (incluso no Python padrão)

### Instalação de dependências

```bash
cd python
pip install -r requirements.txt
```

---

## Como Executar

### Windows (duplo clique)

```
python/executar.bat
```

### Linha de comando

```bash
cd python
python gerador_xml.py
```

---

## Interface

A janela principal é dividida em dois painéis redimensionáveis:

```
┌──────────────────────────────────────────────────────────────────────┐
│  Menu: Arquivo | Ferramentas                                         │
├──────────────────────────────────────────────────────────────────────┤
│  [Selecionar Principal] [Selecionar Origem] [Carregar Planilhas]     │
│  [Copiar Campos] [Validar F5] [Preview F7] [Gerar XMLs F6] [Salvar] │
├────────────────────────────────┬─────────────────────────────────────┤
│  🔍 Filtro por nome/descrição  │  Abas de preview:                   │
│                                │  Validação | LayoutEntrada |        │
│  [Abas da planilha]            │  LayoutPersistencia |               │
│  ┌──────────────────────────┐  │  mapaAtributo | DadoExterno |       │
│  │ Treeview de campos       │  │  ComandoSQL                         │
│  │ (linhas coloridas)       │  │                                     │
│  └──────────────────────────┘  │  [🔄 Atualizar Preview]            │
│                                │                                     │
│  [+Novo][Editar][Remover]      │  (conteúdo XML/SQL/validação)       │
│  [Recalcular]  X/Y campos      │                                     │
├────────────────────────────────┴─────────────────────────────────────┤
│  Status: mensagem dinâmica          Tamanho total: X bytes | Pos: Y  │
└──────────────────────────────────────────────────────────────────────┘
```

### Painel Esquerdo — Tabela de Campos

- **Filtro em tempo real** por nome ou descrição do campo
- **Abas** correspondentes a cada sheet da planilha (ex.: "Campos Entrada")
- **Treeview** com todas as colunas originais da planilha
  - Linhas alternadas (par/ímpar) para legibilidade
  - **Vermelho** — campo com erro de posicionamento
  - **Amarelo** — campo sem posição ou tamanho definido
- **Ações:** `+ Novo`, `✎ Editar` (duplo clique), `🗑 Remover`, `⟳ Recalcular`
- **Contador** e barra de tamanho total ao rodapé

### Painel Direito — Preview e Validação

- **Aba Validação** — resultado colorizado da validação
- **Aba LayoutEntrada** — preview do XML principal
- **Aba LayoutPersistencia** — preview do XML de persistência
- **Aba mapaAtributo** — preview do XML de mapa de atributos
- **Aba DadoExterno** — preview do XML de enriquecimento
- **Aba ComandoSQL** — preview dos scripts SQL

Cada aba possui o botão **🔄 Atualizar Preview** que regenera todas as abas em paralelo (em thread, com loading).

---

## Fluxo de Uso

### 1. Carregar Planilha Principal

1. Clique em **📂 Selecionar Principal** (`Ctrl+O`) e escolha o arquivo `.xlsx`
2. Opcionalmente, clique em **📂 Selecionar Origem** para uma segunda planilha (fonte de cópia)
3. Clique em **⬇ Carregar Planilhas**
   - Uma janela de progresso exibe o arquivo sendo carregado e o tempo decorrido
   - Todas as abas são carregadas e exibidas no painel esquerdo

### 2. Editar Campos

- **Novo campo:** clique em `+ Novo` → preencha o formulário → confirme
- **Editar campo:** duplo clique na linha ou selecione + `✎ Editar`
- **Remover campo:** selecione + `🗑 Remover` → confirme
- **Recalcular posições:** `⟳ Recalcular` redistribui PosIni/PosFin sequencialmente a partir de 1

O campo **PosicaoFinal** é calculado automaticamente (`PosIni + Tamanho - 1`) ao preencher Posição Inicial e Tamanho.

### 3. Copiar Campos da Origem

1. Com a planilha origem carregada, clique em **⬇ Copiar Campos...**
2. Na janela, navegue pelas abas da origem e selecione os campos desejados
   - Seleção múltipla: `Ctrl+Click`, `Shift+Click` ou **Selecionar Todos**
3. Clique em **⬇ Copiar X campos**
   - Campos novos recebem posição e ID sequenciais da planilha principal
   - Campos duplicados: pergunta se deseja atualizar atributos (tipo, tamanho, alinhamento)
   - Dados de persistência e mapa de atributo são mesclados automaticamente da origem
   - O processo roda em thread com barra de progresso e opção de cancelar (com rollback)

### 4. Validar

Clique em **✔ Validar** (`F5`) para verificar:

- Fórmula `PosicaoFinal = PosicaoInicial + TamanhoCampo - 1`
- Ausência de gaps ou sobreposições entre campos
- Início do layout na posição 1

O resultado é exibido na aba **Validação** com erros (vermelho), avisos (laranja) e informações (azul).

### 5. Gerar XMLs

1. Clique em **📄 Gerar XMLs** (`F6`)
2. Se houver erros de validação, o sistema pergunta se deseja continuar
3. Escolha o diretório de saída
4. Os 6 arquivos são gerados em thread com progresso `1/6 … 6/6`
5. O resultado mostra quais arquivos foram gerados com sucesso

### 6. Salvar Planilha

Clique em **💾 Salvar Planilha** (`Ctrl+S`):

- Salva em `{nome_original}_Novo.xlsx` — **o arquivo original nunca é modificado**
- Preserva toda a estrutura: todas as abas, formatação, imagens e seções
- PosicaoFinal é mantida como fórmula Excel (`=L{linha}+K{linha}-1`)

---

## Estrutura da Planilha de Entrada

### Aba principal: `Campos Entrada` (obrigatória)

A aba pode ter uma **linha de metadados** (acima do cabeçalho) indicando seções de colunas, seguida do **cabeçalho** (detectado automaticamente nas primeiras 10 linhas) e dos **dados dos campos**.

```
Linha 1  →  [Metadados de seção]   ex: "Layouts" | "Campos" | "Layout Entrada"
Linha 2  →  [Cabeçalho]            NomeCampo | TipoCampo | TamanhoCampo | ...
Linha 3-5→  (vazias ou reservadas)
Linha 6+ →  [Dados dos campos]
```

#### Colunas reconhecidas

| Coluna | Alternativas aceitas | Descrição |
| --- | --- | --- |
| `Entrada` | — | `S`/`N` — se o campo está ativo no layout de entrada |
| `Persistencia` | — | `S`/`N` — incluir no LayoutPersistencia e SQL |
| `Enriquecimento` | — | `S`/`N` — incluir no DadoExterno |
| `MapaAtributo` | — | `S`/`N` — incluir no mapaAtributo |
| `Saida` | — | `S`/`N` — campo de saída |
| `CampoConcatenado` | — | Campo derivado de concatenação |
| `IdentificadorCampo` | `ID`, `Id` | ID único do campo |
| `NomeCampo` | `Nome`, `Campo` | Nome do campo (**obrigatório**) |
| `DescricaoCampo` | `Descricao` | Descrição do campo |
| `TipoCampo` | `Tipo` | `TEXTO`, `INTEIRO`, `DECIMAL`, `DATA`, `DATA_HORA`, `ID`, `FK`, `NUMERO` |
| `TamanhoCampo` | `Tamanho` | Número de bytes do campo |
| `PosicaoInicial` | `PosInicial`, `PosIni` | Posição inicial no registro |
| `PosicaoFinal` | `PosFinal`, `PosFin` | Posição final (calculada automaticamente) |
| `ValorPadrao` | `Valor_Padrao` | Valor padrão do campo |
| `AlinhamentoCampo` | `Alinhamento` | Tipo de alinhamento e padding |
| `CampoObrigatorio` | `Obrigatorio` | `S`/`N` — campo obrigatório |
| `NomeColuna` | `Coluna_DB`, `ColunaDB` | Nome da coluna no banco de dados |
| `OracleDataType` | `OracleType` | Tipo Oracle (`VARCHAR2`, `NUMBER`, `DATE`) |

### Aba: `Identificação Evento`

Lida automaticamente para extrair metadados usados nos XMLs e no nome dos arquivos gerados.

| Campo | Uso |
| --- | --- |
| `IdentificadorEvento` | Prefixo dos arquivos gerados (`{ID}_Layout_entrada.xml`) |
| `Identificador` | Tag `<Identificador>` no LayoutPersistencia |
| `TamanhoLayout` | Fallback para `<TamanhoLayout>` (padrão: calculado automaticamente) |
| `NomeTabela` | Tag `<NomeTabela>` em todos os `<CampoPersistencia>` |

### Abas de Persistência (`Persistenc*`)

Lidas ao copiar campos com `Persistencia=S` da origem. Colunas mescladas no campo copiado (exceto Persistencia, PosicaoInicial, PosicaoFinal, IdentificadorCampo).

### Abas de Mapa de Atributo (`RuleAttribute*`, `MapaAtributo*`, `AttributeMap*`)

Lidas ao copiar campos com `MapaAtributo=S`. Mesmo comportamento de mesclagem.

### Abas de Enriquecimento

| Aba | Conteúdo |
| --- | --- |
| `Enriquecimento` | Dados principais do DadoAcesso (ComandoSQL, Nome, TamanhoTransacao, etc.) |
| `Enr_ChaveAcesso` | Chaves de acesso linkadas por `IdentificadorEnriquecimento` |
| `Enr_CampoRetornado` | Campos retornados linkados por `IdentificadorEnriquecimento` |

### Aba: `ComandosSQL`

Contém SQL fixos (cabeçalho do script) a inserir antes dos `INSERT`s gerados automaticamente. Linhas onde a coluna 1 = `"insert na tabela column_configuration"` são ignoradas (marcador interno).

### Aba: `Rule Attribute Valor Padrão`

Lida para preencher o bloco `<defaultValueDefinition>` no `mapaAtributo.xml`.

---

## Arquivos Gerados

### `{ID}_Layout_entrada.xml`

Layout posicional completo. Contém todos os campos com `Entrada=S` e posição definida, ordenados por `PosicaoInicial`.

```xml
<?xml version="1.0"?>
<LayoutEntrada>
  <Campos>
    <CampoEntrada>
      <IdentificadorCampo>10</IdentificadorCampo>
      <NomeCampo>CPF_CLIENTE</NomeCampo>
      <DescricaoCampo>CPF do cliente</DescricaoCampo>
      <TipoCampo>TEXTO</TipoCampo>
      <TamanhoCampo>11</TamanhoCampo>
      <AlinhamentoCampo>BRANCO_ESQUERDA</AlinhamentoCampo>
      <Posicao>
        <PosicaoInicial>1</PosicaoInicial>
        <PosicaoFinal>11</PosicaoFinal>
      </Posicao>
    </CampoEntrada>
  </Campos>
</LayoutEntrada>
```

### `{ID}_Layout_persistencia.xml`

Campos com `Persistencia=S`. Metadados de cabeçalho lidos da aba "Identificação Evento". `TamanhoLayout` calculado como `max(PosicaoFinal)` dos Campos Entrada.

```xml
<?xml version="1.0"?>
<LayoutPersistencia>
  <Identificador>1</Identificador>
  <TamanhoLayout>1500</TamanhoLayout>
  <IdentificadorEvento>CLIENTE</IdentificadorEvento>
  <Campos>
    <CampoPersistencia>
      <NomeTabela>TAB_CLIENTE</NomeTabela>
      <NomeColuna>CPF</NomeColuna>
      <AlinhamentoCampo>BRANCO_ESQUERDA</AlinhamentoCampo>
      <IdentificadorCampo>10</IdentificadorCampo>
      <NomeCampo>CPF_CLIENTE</NomeCampo>
      <DescricaoCampo>CPF do cliente</DescricaoCampo>
      <TipoCampo>TEXTO</TipoCampo>
      <CampoObrigatorio>S</CampoObrigatorio>
      <TamanhoCampo>11</TamanhoCampo>
    </CampoPersistencia>
  </Campos>
</LayoutPersistencia>
```

### `{ID}_Layout_mapa_atributo.xml`

Campos com `MapaAtributo=S`. Usa namespace `ns2` (CPQD). Inclui bloco de valores padrão lidos da aba "Rule Attribute Valor Padrão".

```xml
<?xml version="1.0"?>
<ns2:attributeMap xmlns:ns2="http://rule.saf.cpqd.com.br/">
  <defaultValueDefinition>
    <defaultValueItem dataType="STRING" pattern="" value=""/>
  </defaultValueDefinition>
  <input>
    <origin name="ENRICHMENT">
      <attribute>
        <eventAttribute name="CPF_CLIENTE" type="STRING"/>
        <ruleAttribute name="CPF_CLIENTE" type="STRING"/>
        <description>CPF do cliente</description>
        <documentation></documentation>
      </attribute>
    </origin>
  </input>
</ns2:attributeMap>
```

### `{ID}_Layout_enriquecimento.xml`

Gerado a partir das abas `Enriquecimento`, `Enr_ChaveAcesso` e `Enr_CampoRetornado`. Os campos `ComandoSQL` e `SQLChave` são encapsulados em `CDATA`. `TamanhoTransacao` = `max(PosicaoFinal)` dos Campos Entrada.

```xml
<?xml version="1.0" encoding="UTF-8"?>
<DadoExterno>
  <Metrica ligado="S" modo="JMX"/>
  <DadoAcesso>
    <ComandoSQL><![CDATA[SELECT CPF FROM TAB_CLIENTE WHERE ID = ?]]></ComandoSQL>
    <Nome>ENRIQ_CPF</Nome>
    <TamanhoTransacao>1500</TamanhoTransacao>
    <PersistirEnriquecimento>S</PersistirEnriquecimento>
    <GrupoChave>
      <ChaveAcesso>
        <Identificador>1</Identificador>
        <PosInicial>1</PosInicial>
        <PosFinal>11</PosFinal>
      </ChaveAcesso>
    </GrupoChave>
    <CampoRetornado>
      <AliasCampo>CPF</AliasCampo>
      <CampoDestino>CPF_CLIENTE</CampoDestino>
      <TipoCampo>TEXTO</TipoCampo>
      <PosInicial>1</PosInicial>
      <PosFinal>11</PosFinal>
    </CampoRetornado>
  </DadoAcesso>
</DadoExterno>
```

### `ComandoSQL.sql`

Script SQL com os INSERTs para a tabela `COLUMN_CONFIGURATION`. Começa com os SQLs fixos da aba `ComandosSQL` do xlsx, seguido de um INSERT por campo com `Persistencia=S`.

```sql
-- [SQLs fixos da aba ComandosSQL]

insert into COLUMN_CONFIGURATION
(ID_COLUMN_CONFIGURATION,ID_TABLE_CONFIGURATION,ID_DATA_TYPE,
NM_COLUMN_CONFIGURATION,DS_COLUMN_CONFIGURATION,
NR_DATA_LENGTH,NR_DATA_PRECISION,NR_DATA_SCALE,IN_NULLABLE,IN_PK,IN_FK)
values (
  seq_COLUMN_CONFIGURATION.nextval,
  (select ID_TABLE_CONFIGURATION from TABLE_CONFIGURATION
   where NM_TABLE_CONFIGURATION='TAB_CLIENTE'),
  (select ID_DATA_TYPE from DATA_TYPE where NM_DATA_TYPE='VARCHAR2'),
  'CPF','CPF do cliente',11,null,null,1,0,0);
```

**Mapeamento de tipo para SQL:**

| TipoCampo | SQL Type | NR_DATA_LENGTH | NR_DATA_PRECISION |
| --- | --- | --- | --- |
| `TEXTO` | `VARCHAR2` | tamanho | null |
| `INTEIRO`, `ID`, `FK`, `DECIMAL`, `NUMERO`, `NUMBER` | `NUMBER` | null | tamanho |
| `DATA`, `DATA_HORA` | `DATE` | null | null |

### `evento_event_{ID}.xlsx`

Cópia integral da planilha principal com todas as alterações aplicadas. Preserva:

- Todas as abas (inclusive abas não modificadas)
- Formatação, estilos e imagens
- Linhas de metadados e seções acima do cabeçalho
- `PosicaoFinal` como fórmula Excel (`=K{linha}+J{linha}-1`)

---

## Regras de Negócio

### Cálculo de Posições

```
PosicaoFinal = PosicaoInicial + TamanhoCampo - 1
```

Exemplo: `PosIni=10`, `Tamanho=5` → `PosFin=14` (ocupa bytes 10, 11, 12, 13, 14)

### Faixas de ID Reservadas

Ao copiar campos da origem, o sistema atribui IDs sequenciais **pulando automaticamente** as faixas reservadas:

| Faixa | Status |
| --- | --- |
| `1 – 999` | Livre para uso |
| `1000 – 1999` | **Reservada** |
| `2000 – 19999` | Livre para uso |
| `20000 – 21000` | **Reservada** |

### Persistência (`Persistencia=S`)

Quando um campo tem `Persistencia=S`:

- É incluído em `{ID}_Layout_persistencia.xml`
- Gera um INSERT em `ComandoSQL.sql`
- Ao copiar da origem, dados da aba `Persistenc*` são mesclados automaticamente no campo
- `NomeTabela` é sempre forçado para o valor da planilha **principal** (não da origem)

### Mapa de Atributo (`MapaAtributo=S`)

Quando um campo tem `MapaAtributo=S`:

- É incluído em `{ID}_Layout_mapa_atributo.xml`
- Ao copiar da origem, dados das abas `RuleAttribute*` / `MapaAtributo*` são mesclados automaticamente

### Alinhamento de Campos

| Valor | Comportamento |
| --- | --- |
| `BRANCO_ESQUERDA` | Texto alinhado à esquerda, espaços à direita (padrão texto) |
| `BRANCO_DIREITA` | Texto alinhado à direita, espaços à esquerda |
| `ZERO_ESQUERDA` | Número alinhado à direita, zeros à esquerda (padrão numérico) |
| `ZERO_DIREITA` | Número alinhado à esquerda, zeros à direita |

### Salvamento Seguro

- **Nunca sobrescreve** o arquivo original — salva sempre em `{original}_Novo.xlsx`
- Ao gerar XMLs, a cópia da planilha usa `shutil.copy2` (cópia byte-a-byte) e reescreve apenas as células de dados

---

## Validação

O sistema executa as seguintes verificações ao validar (`F5`) ou antes de gerar XMLs (`F6`):

| Verificação | Severidade | Descrição |
| --- | --- | --- |
| Fórmula PosicaoFinal | **ERRO** | `PosIni + Tamanho - 1 ≠ PosFin` |
| Início em 1 | **AVISO** | Primeiro campo não começa na posição 1 |
| Continuidade | **AVISO** | Gap ou sobreposição entre campos consecutivos |
| Campos sem posição | **AVISO** | Campo ativo sem `PosicaoInicial` ou `TamanhoCampo` |

Resultados exibidos na aba **Validação**:

- **Azul** — informações (total de campos, soma de bytes, posição final)
- **Laranja** — avisos (não impedem geração)
- **Vermelho** — erros (pergunta se deseja gerar mesmo assim)

Campos com erro são marcados em **vermelho** na tabela; campos sem posição em **amarelo**.

---

## Atalhos de Teclado

| Atalho | Ação |
| --- | --- |
| `Ctrl+O` | Selecionar planilha principal |
| `Ctrl+S` | Salvar planilha (`_Novo.xlsx`) |
| `F5` | Validar campos |
| `F6` | Gerar todos os XMLs + planilha |
| `F7` | Atualizar todas as abas de preview |
| `Delete` | Remover campo selecionado |
| Duplo clique | Editar campo selecionado |

---

## Janelas de Loading

Todas as operações pesadas rodam em thread separada e exibem uma janela de progresso com:

- Mensagem dinâmica indicando o passo atual
- Barra de progresso indeterminada
- **Timer `MM:SS`** mostrando o tempo decorrido
- **Botão Cancelar** — interrompe o processo e faz rollback automático:
  - **Carregar planilhas** → nenhum dado é aplicado
  - **Copiar campos** → campos já inseridos são removidos
  - **Preview / Gerar XMLs** → nenhuma aba de preview é atualizada

| Operação | Progresso exibido |
| --- | --- |
| Carregar planilhas | `"arquivo 1 de 2"` |
| Copiar campos | `"Copiando campos... X de N"` |
| Atualizar Preview | `"Gerando preview: {aba} — X de 5"` |
| Gerar XMLs | `"Gerando: {arquivo} — X/6"` |

---

## Estrutura do Projeto

```
AlteraEventos/
├── python/
│   ├── gerador_xml.py      # Aplicação principal (Python/Tkinter)
│   ├── requirements.txt    # Dependência: openpyxl>=3.0.10
│   └── executar.bat        # Atalho de execução no Windows
├── src/                    # Código-fonte Java (versão legada)
└── README.md               # Esta documentação
```
