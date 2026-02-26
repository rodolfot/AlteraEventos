package com.alteraeventos.service;

import com.alteraeventos.model.CampoEntrada;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Serviço responsável por ler e escrever dados da planilha Excel.
 * Compatível com Java 8.
 *
 * Estrutura da aba "Campos Entrada":
 *   Linha 1: cabeçalhos de grupo
 *   Linha 2: cabeçalhos das colunas
 *   Linhas 3-5: campos internos do sistema
 *   Linha 6+: campos de entrada do evento
 */
public class PlanilhaService {

    public static final String SHEET_CAMPOS_ENTRADA = "Campos Entrada";

    // Índices das colunas (0-based)
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

    private File arquivoAtual;

    public List<CampoEntrada> lerPlanilha(File arquivo) throws IOException {
        this.arquivoAtual = arquivo;
        List<CampoEntrada> campos = new ArrayList<>();

        FileInputStream fis = new FileInputStream(arquivo);
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(fis);
        } finally {
            fis.close();
        }

        try {
            Sheet sheet = workbook.getSheet(SHEET_CAMPOS_ENTRADA);
            if (sheet == null) {
                throw new IOException("Aba '" + SHEET_CAMPOS_ENTRADA + "' não encontrada.\n"
                        + "Abas disponíveis: " + obterNomesAbas(workbook));
            }

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (int rowIdx = DATA_START_ROW; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row == null) continue;

                String nomeCampo = getCellString(row, COL_NOME_CAMPO, evaluator);
                if (nomeCampo.trim().isEmpty()) continue;

                CampoEntrada campo = new CampoEntrada();
                campo.setLinha(rowIdx + 1);

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

                String valorInicial = campo.getValorPadrao().trim().isEmpty()
                        ? campo.getDefaultValue() : campo.getValorPadrao();
                campo.setValorUsuario(valorInicial);

                campos.add(campo);
            }
        } finally {
            workbook.close();
        }

        return campos;
    }

    public void salvarPlanilha(File arquivo, List<CampoEntrada> campos) throws IOException {
        Workbook workbook;
        FileInputStream fis = new FileInputStream(arquivo);
        try {
            workbook = WorkbookFactory.create(fis);
        } finally {
            fis.close();
        }

        Sheet sheet = workbook.getSheet(SHEET_CAMPOS_ENTRADA);
        if (sheet == null) {
            workbook.close();
            throw new IOException("Aba '" + SHEET_CAMPOS_ENTRADA + "' não encontrada.");
        }

        for (CampoEntrada campo : campos) {
            int rowIdx = campo.getLinha() - 1;
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

        FileOutputStream fos = new FileOutputStream(arquivo);
        try {
            workbook.write(fos);
        } finally {
            fos.close();
            workbook.close();
        }
    }

    public File getArquivoAtual() { return arquivoAtual; }

    // =========================================================
    // Auxiliares de leitura
    // =========================================================

    private String getCellString(Row row, int colIdx, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) return "";
        try {
            CellValue cv = evaluator.evaluate(cell);
            if (cv == null) return "";
            switch (cv.getCellType()) {
                case STRING:
                    return cv.getStringValue().trim();
                case NUMERIC:
                    double d = cv.getNumberValue();
                    if (d == Math.floor(d) && !Double.isInfinite(d)) {
                        return String.valueOf((long) d);
                    }
                    return String.valueOf(d);
                case BOOLEAN:
                    return String.valueOf(cv.getBooleanValue());
                default:
                    return "";
            }
        } catch (Exception e) {
            try { return cell.getStringCellValue().trim(); } catch (Exception ex) { return ""; }
        }
    }

    private Integer getCellInteger(Row row, int colIdx, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;
        try {
            CellValue cv = evaluator.evaluate(cell);
            if (cv == null) return null;
            switch (cv.getCellType()) {
                case NUMERIC:
                    return (int) cv.getNumberValue();
                case STRING:
                    String s = cv.getStringValue().trim();
                    if (s.isEmpty()) return null;
                    try { return Integer.parseInt(s); } catch (NumberFormatException ex) { return null; }
                default:
                    return null;
            }
        } catch (Exception e) {
            return null;
        }
    }

    // =========================================================
    // Auxiliares de escrita
    // =========================================================

    private void setCellString(Row row, int colIdx, String value) {
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx, CellType.STRING);
        } else if (cell.getCellType() == CellType.FORMULA) {
            return; // preserva fórmulas
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
