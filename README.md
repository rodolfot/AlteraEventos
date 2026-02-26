# AlteraEventos — Gerador de XML a partir de Planilhas

Software Java/Swing para leitura de planilhas Excel (.xlsx) com layout de eventos posicionais e geração de XML correspondente.

## Funcionalidades

- **Leitura de Planilha**: carrega arquivos `.xlsx` na aba `Campos Entrada`
- **Interface Gráfica**: tabela editável com todos os campos do evento
- **Validação**: verifica posições, tamanhos e continuidade do layout
- **Atualização da Planilha**: salva alterações de volta no arquivo Excel
- **Geração de XML**: gera XML com atributos posicionados corretamente
- **Pré-visualização**: visualiza o XML antes de salvar

## Estrutura da Planilha Esperada

A planilha deve conter a aba **`Campos Entrada`** com as colunas:

| Coluna | Campo |
|--------|-------|
| A | Entrada (S/N) |
| G | IdentificadorCampo |
| H | NomeCampo |
| I | DescricaoCampo |
| J | TipoCampo |
| K | TamanhoCampo |
| L | PosicaoInicial |
| M | PosicaoFinal (fórmula: =L+K-1) |
| N | ValorPadrao |
| O | AlinhamentoCampo |
| P | CampoObrigatorio |

## Como Executar

### Pré-requisitos
- Java 17+
- Maven 3.8+

### Build
```bash
mvn clean package
```

### Executar
```bash
java -jar target/AlteraEventos.jar
# Ou abrindo diretamente com um arquivo:
java -jar target/AlteraEventos.jar caminho/para/planilha.xlsx
```

## Fluxo de Uso

1. Clique em **Abrir Planilha** e selecione o arquivo `.xlsx`
2. A tabela exibe todos os campos da aba `Campos Entrada`
3. Selecione um campo para editar seus detalhes no painel direito
4. Preencha a coluna **Valor** para geração do XML
5. Clique em **Validar [F5]** para verificar posições e tamanhos
6. Clique em **Atualizar Planilha** para salvar as alterações no Excel
7. Clique em **Gerar XML** (ou **Pré-visualizar XML [F6]**) para o XML final

## Tipos de Alinhamento

| Tipo | Comportamento |
|------|--------------|
| BRANCO_DIREITA | Valor alinhado à direita, espaços à esquerda |
| BRANCO_ESQUERDA | Valor alinhado à esquerda, espaços à direita |
| ZERO_ESQUERDA | Valor alinhado à direita, zeros à esquerda |
| ZERO_DIREITA | Valor alinhado à esquerda, zeros à direita |

## Atalhos de Teclado

| Atalho | Ação |
|--------|------|
| Ctrl+O | Abrir planilha |
| Ctrl+S | Atualizar planilha |
| Ctrl+G | Gerar XML |
| F5 | Validar campos |
| F6 | Pré-visualizar XML |
