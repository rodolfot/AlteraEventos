package com.alteraeventos.ui;

import com.alteraeventos.model.CampoEntrada;
import com.alteraeventos.model.ResultadoValidacao;
import com.alteraeventos.service.ValidacaoService;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.TableRowSorter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.List;
import java.util.regex.PatternSyntaxException;

/**
 * Painel principal de exibição e edição dos campos de entrada.
 * Compatível com Java 8.
 */
public class CamposEntradaPanel extends JPanel {

    private final CamposTableModel tableModel;
    private final JTable tabela;
    private final TableRowSorter<CamposTableModel> sorter;
    private final JTextField filtroField;

    private final ValidacaoService validacaoService = new ValidacaoService();

    // Painel de detalhes
    private final JTextField txtNome = new JTextField();
    private final JTextField txtDescricao = new JTextField();
    private final JComboBox<String> cmbTipo;
    private final JSpinner spnTamanho = new JSpinner(new SpinnerNumberModel(1, 1, 9999, 1));
    private final JSpinner spnPosInicial = new JSpinner(new SpinnerNumberModel(1, 1, 99999, 1));
    private final JTextField txtPosFinal = new JTextField();
    private final JComboBox<String> cmbAlinhamento;
    private final JComboBox<String> cmbObrigatorio;
    private final JComboBox<String> cmbEntrada;
    private final JTextField txtValorPadrao = new JTextField();
    private final JTextField txtValorUsuario = new JTextField();
    private final JTextField txtNomeColuna = new JTextField();
    private final JTextField txtOracleType = new JTextField();

    private Runnable onAtualizarPlanilha;
    private Runnable onGerarXml;

    private int linhaSelecionada = -1;
    private boolean atualizandoDetalhes = false;

    public CamposEntradaPanel() {
        setLayout(new BorderLayout(0, 0));

        cmbTipo = new JComboBox<>(new String[]{
            "TEXTO", "INTEIRO", "INTEIRO_LONGO", "DECIMAL",
            "DATA", "DATA_HORA", "HORA", "BOOLEANO"
        });
        cmbAlinhamento = new JComboBox<>(new String[]{
            "", "BRANCO_ESQUERDA", "BRANCO_DIREITA", "ZERO_ESQUERDA", "ZERO_DIREITA"
        });
        cmbObrigatorio = new JComboBox<>(new String[]{"", "S", "N"});
        cmbEntrada = new JComboBox<>(new String[]{"S", "N"});

        tableModel = new CamposTableModel();
        tabela = new JTable(tableModel);
        sorter = new TableRowSorter<>(tableModel);
        tabela.setRowSorter(sorter);

        configurarTabela();

        JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                criarPainelTabela(), criarPainelDetalhes());
        splitPane.setResizeWeight(0.65);

        add(criarPainelFiltro(), BorderLayout.NORTH);
        add(splitPane, BorderLayout.CENTER);
        add(criarPainelBotoes(), BorderLayout.SOUTH);
    }

    private void configurarTabela() {
        tabela.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        tabela.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        tabela.setFillsViewportHeight(true);
        tabela.setRowHeight(22);
        tabela.getTableHeader().setReorderingAllowed(false);

        int[] larguras = {40, 200, 250, 100, 70, 85, 75, 140, 60, 150};
        for (int i = 0; i < larguras.length && i < tabela.getColumnCount(); i++) {
            tabela.getColumnModel().getColumn(i).setPreferredWidth(larguras[i]);
        }

        tabela.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent e) {
                if (!e.getValueIsAdjusting()) {
                    int viewRow = tabela.getSelectedRow();
                    if (viewRow >= 0) {
                        linhaSelecionada = tabela.convertRowIndexToModel(viewRow);
                        preencherDetalhes(tableModel.getCampo(linhaSelecionada));
                    } else {
                        linhaSelecionada = -1;
                        limparDetalhes();
                    }
                }
            }
        });

        tabela.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) {
                    txtValorUsuario.requestFocusInWindow();
                    txtValorUsuario.selectAll();
                }
            }
        });
    }

    private JPanel criarPainelFiltro() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 4));
        panel.setBorder(BorderFactory.createEmptyBorder(4, 4, 0, 4));

        filtroField = new JTextField(30);
        filtroField.setToolTipText("Filtrar por nome ou descrição");

        filtroField.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent e) { aplicarFiltro(); }
            public void removeUpdate(DocumentEvent e) { aplicarFiltro(); }
            public void changedUpdate(DocumentEvent e) { aplicarFiltro(); }
        });

        JButton btnLimpar = new JButton("Limpar");
        btnLimpar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                filtroField.setText("");
                aplicarFiltro();
            }
        });

        panel.add(new JLabel("Filtrar:"));
        panel.add(filtroField);
        panel.add(btnLimpar);
        return panel;
    }

    private JPanel criarPainelTabela() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEtchedBorder(), "Campos de Entrada",
                TitledBorder.LEFT, TitledBorder.TOP));
        panel.add(new JScrollPane(tabela), BorderLayout.CENTER);
        return panel;
    }

    private JPanel criarPainelDetalhes() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEtchedBorder(), "Detalhes do Campo",
                TitledBorder.LEFT, TitledBorder.TOP));
        panel.setMinimumSize(new Dimension(280, 0));

        txtPosFinal.setEditable(false);
        txtPosFinal.setBackground(UIManager.getColor("TextField.disabledBackground"));
        txtNomeColuna.setEditable(false);
        txtNomeColuna.setBackground(UIManager.getColor("TextField.disabledBackground"));
        txtOracleType.setEditable(false);
        txtOracleType.setBackground(UIManager.getColor("TextField.disabledBackground"));

        JPanel form = new JPanel(new GridBagLayout());
        form.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
        GridBagConstraints lb = new GridBagConstraints();
        lb.anchor = GridBagConstraints.WEST;
        lb.insets = new Insets(3, 0, 3, 6);
        lb.gridx = 0; lb.weightx = 0;

        GridBagConstraints fb = new GridBagConstraints();
        fb.fill = GridBagConstraints.HORIZONTAL;
        fb.insets = new Insets(3, 0, 3, 0);
        fb.gridx = 1; fb.weightx = 1.0;

        int row = 0;
        addRow(form, lb, fb, row++, "Entrada:", cmbEntrada);
        addRow(form, lb, fb, row++, "Nome:", txtNome);
        addRow(form, lb, fb, row++, "Descrição:", txtDescricao);
        addRow(form, lb, fb, row++, "Tipo:", cmbTipo);
        addRow(form, lb, fb, row++, "Tamanho:", spnTamanho);
        addRow(form, lb, fb, row++, "Pos. Inicial:", spnPosInicial);
        addRow(form, lb, fb, row++, "Pos. Final:", txtPosFinal);
        addRow(form, lb, fb, row++, "Alinhamento:", cmbAlinhamento);
        addRow(form, lb, fb, row++, "Obrigatório:", cmbObrigatorio);
        addRow(form, lb, fb, row++, "Valor Padrão:", txtValorPadrao);

        GridBagConstraints sepC = new GridBagConstraints();
        sepC.gridx = 0; sepC.gridy = row++; sepC.gridwidth = 2;
        sepC.fill = GridBagConstraints.HORIZONTAL;
        sepC.insets = new Insets(6, 0, 6, 0);
        form.add(new JSeparator(), sepC);

        addRow(form, lb, fb, row++, "Valor (XML):", txtValorUsuario);
        addRow(form, lb, fb, row++, "Coluna DB:", txtNomeColuna);
        addRow(form, lb, fb, row, "Tipo Oracle:", txtOracleType);

        // Listeners para recalcular Pos. Final automaticamente
        ChangeListener posListener = new ChangeListener() {
            public void stateChanged(ChangeEvent e) { atualizarPosFinal(); }
        };
        spnTamanho.addChangeListener(posListener);
        spnPosInicial.addChangeListener(posListener);

        JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 4, 4));
        JButton btnAplicar = new JButton("Aplicar Alterações");
        btnAplicar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { aplicarAlteracoesDetalhes(); }
        });
        btnPanel.add(btnAplicar);

        panel.add(new JScrollPane(form), BorderLayout.CENTER);
        panel.add(btnPanel, BorderLayout.SOUTH);

        return panel;
    }

    private void addRow(JPanel panel, GridBagConstraints lb, GridBagConstraints fb,
                         int row, String label, JComponent comp) {
        lb.gridy = row;
        fb.gridy = row;
        panel.add(new JLabel(label), lb);
        panel.add(comp, fb);
    }

    private JPanel criarPainelBotoes() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 6));
        panel.setBorder(BorderFactory.createEmptyBorder(0, 4, 4, 4));

        JButton btnValidar = new JButton("Validar Campos");
        btnValidar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { validarCampos(); }
        });

        JButton btnRecalcular = new JButton("Recalcular Posições");
        btnRecalcular.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { recalcularPosicoes(); }
        });

        JButton btnAtualizar = new JButton("Atualizar Planilha");
        btnAtualizar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                if (onAtualizarPlanilha != null) onAtualizarPlanilha.run();
            }
        });

        JButton btnGerarXml = new JButton("Gerar XML");
        btnGerarXml.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                if (onGerarXml != null) onGerarXml.run();
            }
        });

        panel.add(btnValidar);
        panel.add(btnRecalcular);
        panel.add(new JSeparator(SwingConstants.VERTICAL));
        panel.add(btnAtualizar);
        panel.add(btnGerarXml);
        return panel;
    }

    // =========================================================
    // Detalhes
    // =========================================================

    private void atualizarPosFinal() {
        if (atualizandoDetalhes) return;
        int posInicial = (Integer) spnPosInicial.getValue();
        int tamanho = (Integer) spnTamanho.getValue();
        txtPosFinal.setText(String.valueOf(posInicial + tamanho - 1));
    }

    private void preencherDetalhes(CampoEntrada campo) {
        atualizandoDetalhes = true;
        try {
            cmbEntrada.setSelectedItem(campo.getEntrada().trim().isEmpty() ? "S" : campo.getEntrada());
            txtNome.setText(campo.getNomeCampo());
            txtDescricao.setText(campo.getDescricaoCampo());
            cmbTipo.setSelectedItem(campo.getTipoCampo().trim().isEmpty() ? "TEXTO" : campo.getTipoCampo());

            int tamanho = campo.getTamanhoCampo() != null ? campo.getTamanhoCampo() : 1;
            int posInicial = campo.getPosicaoInicial() != null ? campo.getPosicaoInicial() : 1;

            spnTamanho.setValue(tamanho);
            spnPosInicial.setValue(posInicial);
            txtPosFinal.setText(String.valueOf(posInicial + tamanho - 1));

            cmbAlinhamento.setSelectedItem(campo.getAlinhamentoCampo().trim().isEmpty()
                    ? "" : campo.getAlinhamentoCampo());
            cmbObrigatorio.setSelectedItem(campo.getCampoObrigatorio().trim().isEmpty()
                    ? "" : campo.getCampoObrigatorio());

            txtValorPadrao.setText(campo.getValorPadrao());
            txtValorUsuario.setText(campo.getValorUsuario());
            txtNomeColuna.setText(campo.getNomeColuna());
            txtOracleType.setText(campo.getOracleDataType());
        } finally {
            atualizandoDetalhes = false;
        }
    }

    private void limparDetalhes() {
        atualizandoDetalhes = true;
        try {
            cmbEntrada.setSelectedIndex(0);
            txtNome.setText("");
            txtDescricao.setText("");
            cmbTipo.setSelectedIndex(0);
            spnTamanho.setValue(1);
            spnPosInicial.setValue(1);
            txtPosFinal.setText("");
            cmbAlinhamento.setSelectedIndex(0);
            cmbObrigatorio.setSelectedIndex(0);
            txtValorPadrao.setText("");
            txtValorUsuario.setText("");
            txtNomeColuna.setText("");
            txtOracleType.setText("");
        } finally {
            atualizandoDetalhes = false;
        }
    }

    private void aplicarAlteracoesDetalhes() {
        if (linhaSelecionada < 0) {
            JOptionPane.showMessageDialog(this, "Selecione um campo na tabela para editar.",
                    "Nenhum campo selecionado", JOptionPane.WARNING_MESSAGE);
            return;
        }

        CampoEntrada campo = tableModel.getCampo(linhaSelecionada);
        campo.setEntrada((String) cmbEntrada.getSelectedItem());
        campo.setNomeCampo(txtNome.getText().trim());
        campo.setDescricaoCampo(txtDescricao.getText().trim());
        campo.setTipoCampo((String) cmbTipo.getSelectedItem());
        campo.setTamanhoCampo((Integer) spnTamanho.getValue());
        campo.setPosicaoInicial((Integer) spnPosInicial.getValue());
        campo.setPosicaoFinal((Integer) spnPosInicial.getValue() + (Integer) spnTamanho.getValue() - 1);
        campo.setAlinhamentoCampo((String) cmbAlinhamento.getSelectedItem());
        campo.setCampoObrigatorio((String) cmbObrigatorio.getSelectedItem());
        campo.setValorPadrao(txtValorPadrao.getText().trim());
        campo.setValorUsuario(txtValorUsuario.getText());

        tableModel.atualizarCampo(linhaSelecionada, campo);

        JOptionPane.showMessageDialog(this,
                "Alterações aplicadas para '" + campo.getNomeCampo() + "'.\n"
                + "Clique em 'Atualizar Planilha' para salvar no Excel.",
                "Alterações aplicadas", JOptionPane.INFORMATION_MESSAGE);
    }

    // =========================================================
    // Ações dos botões
    // =========================================================

    private void aplicarFiltro() {
        String texto = filtroField.getText().trim();
        if (texto.isEmpty()) { sorter.setRowFilter(null); return; }
        try {
            sorter.setRowFilter(RowFilter.regexFilter("(?i)" + texto, 1, 2));
        } catch (PatternSyntaxException e) {
            sorter.setRowFilter(null);
        }
    }

    private void validarCampos() {
        List<CampoEntrada> campos = tableModel.getCampos();
        if (campos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Nenhum campo carregado.",
                    "Validação", JOptionPane.WARNING_MESSAGE);
            return;
        }

        ResultadoValidacao resultado = validacaoService.validar(campos);

        JTextArea textArea = new JTextArea(resultado.getRelatorio(), 20, 60);
        textArea.setEditable(false);
        textArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        JScrollPane scroll = new JScrollPane(textArea);

        JOptionPane.showMessageDialog(this, scroll,
                resultado.isValido() ? "Validação OK" : "Validação com Problemas",
                resultado.isValido() ? JOptionPane.INFORMATION_MESSAGE : JOptionPane.WARNING_MESSAGE);
    }

    private void recalcularPosicoes() {
        int opcao = JOptionPane.showConfirmDialog(this,
                "Isso vai recalcular posições iniciais e finais de todos os campos\n"
                + "em sequência começando de 1. Deseja continuar?",
                "Recalcular Posições", JOptionPane.YES_NO_OPTION);
        if (opcao == JOptionPane.YES_OPTION) {
            validacaoService.recalcularPosicoes(tableModel.getCampos());
            tableModel.fireTableDataChanged();
            JOptionPane.showMessageDialog(this,
                    "Posições recalculadas. Clique em 'Atualizar Planilha' para salvar.",
                    "Recalcular Posições", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    // =========================================================
    // API Pública
    // =========================================================

    public void carregarCampos(List<CampoEntrada> campos) {
        tableModel.setCampos(campos);
        linhaSelecionada = -1;
        limparDetalhes();
        filtroField.setText("");
        sorter.setRowFilter(null);
    }

    public List<CampoEntrada> getCampos() { return tableModel.getCampos(); }

    public void setOnAtualizarPlanilha(Runnable callback) { this.onAtualizarPlanilha = callback; }
    public void setOnGerarXml(Runnable callback) { this.onGerarXml = callback; }
}
