package com.alteraeventos.ui;

import com.alteraeventos.model.CampoEntrada;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.List;

/**
 * TableModel para exibir os campos de entrada em uma JTable.
 * Permite edição do campo "Valor" (coluna 9).
 */
public class CamposTableModel extends AbstractTableModel {

    private static final String[] NOMES_COLUNAS = {
        "ID", "Nome do Campo", "Descrição", "Tipo", "Tamanho",
        "Pos. Inicial", "Pos. Final", "Alinhamento", "Obrig.", "Valor"
    };

    private static final Class<?>[] TIPOS_COLUNAS = {
        Integer.class, String.class, String.class, String.class, Integer.class,
        Integer.class, Integer.class, String.class, String.class, String.class
    };

    private List<CampoEntrada> campos;

    public CamposTableModel() {
        this.campos = new ArrayList<>();
    }

    public CamposTableModel(List<CampoEntrada> campos) {
        this.campos = new ArrayList<>(campos);
    }

    public void setCampos(List<CampoEntrada> campos) {
        this.campos = new ArrayList<>(campos);
        fireTableDataChanged();
    }

    public List<CampoEntrada> getCampos() {
        return campos;
    }

    public CampoEntrada getCampo(int rowIndex) {
        return campos.get(rowIndex);
    }

    public void adicionarCampo(CampoEntrada campo) {
        campos.add(campo);
        fireTableRowsInserted(campos.size() - 1, campos.size() - 1);
    }

    public void removerCampo(int rowIndex) {
        campos.remove(rowIndex);
        fireTableRowsDeleted(rowIndex, rowIndex);
    }

    public void atualizarCampo(int rowIndex, CampoEntrada campo) {
        campos.set(rowIndex, campo);
        fireTableRowsUpdated(rowIndex, rowIndex);
    }

    @Override
    public int getRowCount() {
        return campos != null ? campos.size() : 0;
    }

    @Override
    public int getColumnCount() {
        return NOMES_COLUNAS.length;
    }

    @Override
    public String getColumnName(int col) {
        return NOMES_COLUNAS[col];
    }

    @Override
    public Class<?> getColumnClass(int col) {
        return TIPOS_COLUNAS[col];
    }

    @Override
    public boolean isCellEditable(int row, int col) {
        // Apenas a coluna "Valor" é editável diretamente na tabela
        return col == 9;
    }

    @Override
    public Object getValueAt(int row, int col) {
        CampoEntrada c = campos.get(row);
        return switch (col) {
            case 0 -> c.getIdentificadorCampo();
            case 1 -> c.getNomeCampo();
            case 2 -> c.getDescricaoCampo();
            case 3 -> c.getTipoCampo();
            case 4 -> c.getTamanhoCampo();
            case 5 -> c.getPosicaoInicial();
            case 6 -> c.getPosicaoFinal() != null
                    ? c.getPosicaoFinal()
                    : c.calcularPosicaoFinalEsperada();
            case 7 -> c.getAlinhamentoCampo();
            case 8 -> c.getCampoObrigatorio();
            case 9 -> c.getValorUsuario();
            default -> null;
        };
    }

    @Override
    public void setValueAt(Object value, int row, int col) {
        if (col == 9) {
            campos.get(row).setValorUsuario(value != null ? value.toString() : "");
            fireTableCellUpdated(row, col);
        }
    }
}
