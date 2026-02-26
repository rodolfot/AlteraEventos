package com.alteraeventos.service;

import com.alteraeventos.model.CampoEntrada;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Serviço responsável por ler e escrever dados da planilha Excel.
 * Suporta a aba "Campos Entrada" conforme especificação.
 *
 * Estrutura da aba "Campos Entrada":
 *   Linha 1: cabeçalhos de grupo
 *   Linha 2: cabeçalhos das colunas
 *   Linhas 3-5: campos internos do sistema (sem posição de layout)
 *   Linha 6+: campos de entrada do evento
 */
public class PlanilhaService {

    public static final String SHEET_CAMPOS_ENTRADA = "Campos Entrada";
    public static final String SHEET_CHAMADA = "Chamada";

    // Índices das colunas na aba "Campos Entrada" (0-based)
    private static final int COL_ENTRADA            = 0;
    private static final int COL_PERSISTENCIA       = 1;
    private static final int COL_ENRIQUECIMENTO     = 2;
    private static final int COL_MAPA_ATRIBUTO      = 3;
    private static final int COL_SAIDA              = 4;
    private static final int COL_CAMPO_CONCATENADO  = 5;
    private static final int COL_ID_CAMPO           = 6;
    private static final int COL_NOME_CAMPO         = 7;
    private static final int COL_DESCRICAO          = 8;
    private static final int COL_TIPO               = 9;
    private static final int COL_TAMANHO            = 10;
    private static final int COL_POS_INICIAL        = 11;
    private static final int COL_POS_FINAL          = 12;
    private static final int COL_VALOR_PADRAO       = 13;
    private static final int COL_ALINHAMENTO        = 14;
    private static final int COL_OBRIGATORIO        = 15;
    private static final int COL_DOMINIO            = 16;
    private static final int COL_MASCARA            = 17;
    private static final int COL_NOME_TABELA        = 22;
    private static final int COL_NOME_COLUNA        = 23;
    private static final int COL_ORACLE_TYPE        = 24;
    private static final int COL_DATA_LENGTH        = 25;
    private static final int COL_NUMBER_PRECISION   = 26;
    private static final int COL_NUMBER_SCALE       = 27;
    private static final int COL_NULLABLE           = 28;
    private static final int COL_ENCRYPTED          = 29;
    private static final int COL_UNIQUE             = 30;
    private static final int COL_RULE_ATTRIBUTE     = 31;
    private static final int COL_DEFAULT_VALUE      = 32;
    private static final int COL_DESCRIPTION        = 33;
    private static final int COL_ORIGIN             = 35;
    private static final int COL_EVENT_ATTRIBUTE    = 37;
    private static final int COL_TYPE               = 38;
    private static final int COL_MODEL_ATTRIBUTE    = 39;
    private static final int COL_SCORE_MODEL_IN     = 40;

    // Linha de início dos dados (0-based) = linha 6 da planilha
    private static final int DATA_START_ROW = 5;

    // Arquivo carregado atualmente
    private File arquivoAtual;

    /**
     * Lê todos os campos de entrada da planilha Excel.
     */
    public List<CampoEntrada> lerPlanilha(File arquivo) throws IOException {
        this.arquivoAtual = arquivo;
        List<CampoEntrada> campos = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(arquivo);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheet(SHEET_CAMPOS_ENTRADA);
            if (sheet == null) {
                throw new IOException("Aba '" + SHEET_CAMPOS_ENTRADA + "' não encontrada na planilha.\n"
                        + "Abas disponíveis: " + obterNomesAbas(workbook));
            }

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (int rowIdx = DATA_START_ROW; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) continue;

                String nomeCampo = getCellString(row, COL_NOME_CAMPO, evaluator);
                if (nomeCampo.isBlank()) continue;

                CampoEntrada campo = new CampoEntrada();
                campo.setLinha(rowIdx + 1); // converte para 1-based

                campo.setEntrada(getCellString(row, COL_ENTRADA, evaluator));
                campo.setPersistencia(getCellString(row, COL_PERSISTENCIA, evaluator));
                campo.setEnriquecimento(getCellString(row, COL_ENRIQUECIMENTO, evaluator));
                campo.setMapaAtributo(getCellString(row, COL_MAPA_ATRIBUTO, evaluator));
                campo.setSaida(getCellString(row, COL_SAIDA, evaluator));
                campo.setCampoConcatenado(getCellString(row, COL_CAMPO_CONCATENADO, evaluator));
                campo.setIdentificadorCampo(getCellInteger(row, COL_ID_CAMPO, evaluator));
                campo.setNomeCampo(nomeCampo);
                campo.setDescricaoCampo(getCellString(row, COL_DESCRICAO, evaluator));
                campo.setTipoCampo(getCellString(row, COL_TIPO, evaluator));
                campo.setTamanhoCampo(getCellInteger(row, COL_TAMANHO, evaluator));
                campo.setPosicaoInicial(getCellInteger(row, COL_POS_INICIAL, evaluator));
                campo.setPosicaoFinal(getCellInteger(row, COL_POS_FINAL, evaluator));
                campo.setValorPadrao(getCellString(row, COL_VALOR_PADRAO, evaluator));
                campo.setAlinhamentoCampo(getCellString(row, COL_ALINHAMENTO, evaluator));
                campo.setCampoObrigatorio(getCellString(row, COL_OBRIGATORIO, evaluator));
                campo.setDominioCampo(getCellString(row, COL_DOMINIO, evaluator));
                campo.setMascaraCampo(getCellString(row, COL_MASCARA, evaluator));
                campo.setNomeTabela(getCellString(row, COL_NOME_TABELA, evaluator));
                campo.setNomeColuna(getCellString(row, COL_NOME_COLUNA, evaluator));
                campo.setOracleDataType(getCellString(row, COL_ORACLE_TYPE, evaluator));
                campo.setDataLength(getCellInteger(row, COL_DATA_LENGTH, evaluator));
                campo.setNumberPrecision(getCellInteger(row, COL_NUMBER_PRECISION, evaluator));
                campo.setNumberScale(getCellInteger(row, COL_NUMBER_SCALE, evaluator));
                campo.setNullable(getCellString(row, COL_NULLABLE, evaluator));
                campo.setEncrypted(getCellString(row, COL_ENCRYPTED, evaluator));
                campo.setUnique(getCellString(row, COL_UNIQUE, evaluator));
                campo.setRuleAttribute(getCellString(row, COL_RULE_ATTRIBUTE, evaluator));
                campo.setDefaultValue(getCellString(row, COL_DEFAULT_VALUE, evaluator));
                campo.setDescription(getCellString(row, COL_DESCRIPTION, evaluator));
                campo.setOrigin(getCellString(row, COL_ORIGIN, evaluator));
                campo.setEventAttribute(getCellString(row, COL_EVENT_ATTRIBUTE, evaluator));
                campo.setType(getCellString(row, COL_TYPE, evaluator));
                campo.setModelAttribute(getCellString(row, COL_MODEL_ATTRIBUTE, evaluator));
                campo.setScoreModelIn(getCellString(row, COL_SCORE_MODEL_IN, evaluator));

                // Pré-preenche o valor com o valor padrão
                String valorInicial = campo.getValorPadrao().isBlank()
                        ? campo.getDefaultValue()
                        : campo.getValorPadrao();
                campo.setValorUsuario(valorInicial);

                campos.add(campo);
            }
        }

        return campos;
    }

    /**
     * Salva as alterações da lista de campos de volta na planilha.
     * Preserva fórmulas existentes (como PosicaoFinal) e estilos.
     */
    public void salvarPlanilha(File arquivo, List<CampoEntrada> campos) throws IOException {
        Workbook workbook;
        try (FileInputStream fis = new FileInputStream(arquivo)) {
            workbook = WorkbookFactory.create(fis);
        }

        Sheet sheet = workbook.getSheet(SHEET_CAMPOS_ENTRADA);
        if (sheet == null) {
            workbook.close();
            throw new IOException("Aba '" + SHEET_CAMPOS_ENTRADA + "' não encontrada.");
        }

        for (CampoEntrada campo : campos) {
            int rowIdx = campo.getLinha() - 1; // converte para 0-based
            Row row = sheet.getRow(rowIdx);
            if (row == null) {
                row = sheet.createRow(rowIdx);
            }

            setCellString(row, COL_NOME_CAMPO, campo.getNomeCampo());
            setCellString(row, COL_DESCRICAO, campo.getDescricaoCampo());
            setCellString(row, COL_TIPO, campo.getTipoCampo());
            setCellString(row, COL_ALINHAMENTO, campo.getAlinhamentoCampo());
            setCellString(row, COL_OBRIGATORIO, campo.getCampoObrigatorio());
            setCellString(row, COL_VALOR_PADRAO, campo.getValorPadrao());
            setCellString(row, COL_ENTRADA, campo.getEntrada());
            setCellString(row, COL_PERSISTENCIA, campo.getPersistencia());
            setCellString(row, COL_ENRIQUECIMENTO, campo.getEnriquecimento());
            setCellString(row, COL_SAIDA, campo.getSaida());

            if (campo.getTamanhoCampo() != null) {
                setCellNumeric(row, COL_TAMANHO, campo.getTamanhoCampo());
            }
            if (campo.getPosicaoInicial() != null) {
                setCellNumeric(row, COL_POS_INICIAL, campo.getPosicaoInicial());
            }
        }

        try (FileOutputStream fos = new FileOutputStream(arquivo)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    /**
     * Adiciona uma nova linha de campo na planilha.
     */
    public void adicionarCampo(File arquivo, CampoEntrada campo) throws IOException {
        Workbook workbook;
        try (FileInputStream fis = new FileInputStream(arquivo)) {
            workbook = WorkbookFactory.create(fis);
        }

        Sheet sheet = workbook.getSheet(SHEET_CAMPOS_ENTRADA);
        if (sheet == null) {
            workbook.close();
            throw new IOException("Aba '" + SHEET_CAMPOS_ENTRADA + "' não encontrada.");
        }

        // Encontra próxima linha disponível
        int nextRow = sheet.getLastRowNum() + 1;
        Row row = sheet.createRow(nextRow);
        campo.setLinha(nextRow + 1);

        setCellString(row, COL_ENTRADA, campo.getEntrada());
        setCellString(row, COL_NOME_CAMPO, campo.getNomeCampo());
        setCellString(row, COL_DESCRICAO, campo.getDescricaoCampo());
        setCellString(row, COL_TIPO, campo.getTipoCampo());
        setCellString(row, COL_ALINHAMENTO, campo.getAlinhamentoCampo());
        setCellString(row, COL_OBRIGATORIO, campo.getCampoObrigatorio());
        setCellString(row, COL_VALOR_PADRAO, campo.getValorPadrao());

        if (campo.getTamanhoCampo() != null) {
            setCellNumeric(row, COL_TAMANHO, campo.getTamanhoCampo());
        }
        if (campo.getPosicaoInicial() != null) {
            setCellNumeric(row, COL_POS_INICIAL, campo.getPosicaoInicial());
            // Adiciona fórmula para PosicaoFinal
            Cell cellPosFinal = row.createCell(COL_POS_FINAL, CellType.FORMULA);
            cellPosFinal.setCellFormula(
                    "L" + (nextRow + 1) + "+K" + (nextRow + 1) + "-1"
            );
        }

        try (FileOutputStream fos = new FileOutputStream(arquivo)) {
            workbook.write(fos);
        }
        workbook.close();
    }

    public File getArquivoAtual() { return arquivoAtual; }

    // =========================================================
    // Métodos auxiliares de leitura de células
    // =========================================================

    private String getCellString(Row row, int colIdx, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) return "";
        try {
            CellValue cv = evaluator.evaluate(cell);
            if (cv == null) return "";
            return switch (cv.getCellType()) {
                case STRING -> cv.getStringValue().trim();
                case NUMERIC -> {
                    double d = cv.getNumberValue();
                    yield (d == Math.floor(d) && !Double.isInfinite(d))
                            ? String.valueOf((long) d)
                            : String.valueOf(d);
                }
                case BOOLEAN -> String.valueOf(cv.getBooleanValue());
                default -> "";
            };
        } catch (Exception e) {
            // Fallback: lê como string diretamente
            try {
                return cell.getStringCellValue().trim();
            } catch (Exception ex) {
                return "";
            }
        }
    }

    private Integer getCellInteger(Row row, int colIdx, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;
        try {
            CellValue cv = evaluator.evaluate(cell);
            if (cv == null) return null;
            return switch (cv.getCellType()) {
                case NUMERIC -> (int) cv.getNumberValue();
                case STRING -> {
                    String s = cv.getStringValue().trim();
                    yield s.isBlank() ? null : Integer.parseInt(s);
                }
                default -> null;
            };
        } catch (Exception e) {
            return null;
        }
    }

    // =========================================================
    // Métodos auxiliares de escrita de células
    // =========================================================

    private void setCellString(Row row, int colIdx, String value) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx, CellType.STRING);
        } else if (cell.getCellType() == CellType.FORMULA) {
            // Não sobrescreve fórmulas
            return;
        }
        cell.setCellValue(value != null ? value : "");
    }

    private void setCellNumeric(Row row, int colIdx, double value) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx, CellType.NUMERIC);
        } else if (cell.getCellType() == CellType.FORMULA) {
            return;
        }
        cell.setCellValue(value);
    }

    private String obterNomesAbas(Workbook wb) {
        List<String> nomes = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            nomes.add(wb.getSheetName(i));
        }
        return String.join(", ", nomes);
    }
}
