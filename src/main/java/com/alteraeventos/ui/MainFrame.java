package com.alteraeventos.ui;

import com.alteraeventos.model.CampoEntrada;
import com.alteraeventos.model.ResultadoValidacao;
import com.alteraeventos.service.PlanilhaService;
import com.alteraeventos.service.ValidacaoService;
import com.alteraeventos.service.XmlService;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.util.List;
import java.util.prefs.Preferences;

/**
 * Janela principal do sistema AlteraEventos.
 *
 * Fluxo:
 *   1. Usuário abre a planilha (.xlsx)
 *   2. Software exibe os campos da aba "Campos Entrada"
 *   3. Usuário edita valores e configurações
 *   4. Usuário valida, atualiza planilha e/ou gera XML
 */
public class MainFrame extends JFrame {

    private static final String PREF_ULTIMO_DIRETORIO = "ultimoDiretorio";
    private static final String TITULO = "AlteraEventos - Gerador de XML a partir de Planilhas";

    private final PlanilhaService planilhaService = new PlanilhaService();
    private final ValidacaoService validacaoService = new ValidacaoService();
    private final XmlService xmlService = new XmlService();
    private final Preferences prefs = Preferences.userNodeForPackage(MainFrame.class);

    private final CamposEntradaPanel camposPanel;
    private final JLabel statusLabel;
    private final JLabel arquivoLabel;

    private File arquivoAtual;
    private boolean alteracoesPendentes = false;

    public MainFrame() {
        setTitle(TITULO);
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        setMinimumSize(new Dimension(1100, 700));
        setPreferredSize(new Dimension(1300, 800));

        // Inicializa os painéis
        camposPanel = new CamposEntradaPanel();
        camposPanel.setOnAtualizarPlanilha(this::atualizarPlanilha);
        camposPanel.setOnGerarXml(this::gerarXml);

        // Status bar
        statusLabel = new JLabel("Pronto. Abra uma planilha para começar.");
        statusLabel.setBorder(BorderFactory.createEmptyBorder(2, 8, 2, 8));

        arquivoLabel = new JLabel("Nenhum arquivo carregado");
        arquivoLabel.setForeground(Color.GRAY);
        arquivoLabel.setBorder(BorderFactory.createEmptyBorder(2, 8, 2, 8));

        // Layout
        setJMenuBar(criarMenuBar());
        add(criarToolBar(), BorderLayout.NORTH);
        add(criarPainelPrincipal(), BorderLayout.CENTER);
        add(criarStatusBar(), BorderLayout.SOUTH);

        // Confirmação ao fechar
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                fecharAplicacao();
            }
        });

        pack();
        setLocationRelativeTo(null);
    }

    // =========================================================
    // Construção da Interface
    // =========================================================

    private JMenuBar criarMenuBar() {
        JMenuBar menuBar = new JMenuBar();

        // Menu Arquivo
        JMenu menuArquivo = new JMenu("Arquivo");
        menuArquivo.setMnemonic('A');

        JMenuItem itemAbrir = new JMenuItem("Abrir Planilha...");
        itemAbrir.setAccelerator(KeyStroke.getKeyStroke("control O"));
        itemAbrir.addActionListener(e -> abrirPlanilha());

        JMenuItem itemSalvar = new JMenuItem("Atualizar Planilha");
        itemSalvar.setAccelerator(KeyStroke.getKeyStroke("control S"));
        itemSalvar.addActionListener(e -> atualizarPlanilha());

        JMenuItem itemGerarXml = new JMenuItem("Gerar XML...");
        itemGerarXml.setAccelerator(KeyStroke.getKeyStroke("control G"));
        itemGerarXml.addActionListener(e -> gerarXml());

        JMenuItem itemSair = new JMenuItem("Sair");
        itemSair.setAccelerator(KeyStroke.getKeyStroke("alt F4"));
        itemSair.addActionListener(e -> fecharAplicacao());

        menuArquivo.add(itemAbrir);
        menuArquivo.add(itemSalvar);
        menuArquivo.addSeparator();
        menuArquivo.add(itemGerarXml);
        menuArquivo.addSeparator();
        menuArquivo.add(itemSair);

        // Menu Ferramentas
        JMenu menuFerramentas = new JMenu("Ferramentas");
        menuFerramentas.setMnemonic('F');

        JMenuItem itemValidar = new JMenuItem("Validar Campos");
        itemValidar.setAccelerator(KeyStroke.getKeyStroke("F5"));
        itemValidar.addActionListener(e -> validarCampos());

        JMenuItem itemPrevisualizar = new JMenuItem("Pré-visualizar XML");
        itemPrevisualizar.setAccelerator(KeyStroke.getKeyStroke("F6"));
        itemPrevisualizar.addActionListener(e -> previsualizarXml());

        menuFerramentas.add(itemValidar);
        menuFerramentas.add(itemPrevisualizar);

        // Menu Ajuda
        JMenu menuAjuda = new JMenu("Ajuda");
        menuAjuda.setMnemonic('j');

        JMenuItem itemSobre = new JMenuItem("Sobre");
        itemSobre.addActionListener(e -> mostrarSobre());

        menuAjuda.add(itemSobre);

        menuBar.add(menuArquivo);
        menuBar.add(menuFerramentas);
        menuBar.add(menuAjuda);

        return menuBar;
    }

    private JToolBar criarToolBar() {
        JToolBar toolBar = new JToolBar();
        toolBar.setFloatable(false);

        JButton btnAbrir = new JButton("Abrir Planilha");
        btnAbrir.setToolTipText("Abrir arquivo Excel (.xlsx)");
        btnAbrir.addActionListener(e -> abrirPlanilha());

        JButton btnSalvar = new JButton("Atualizar Planilha");
        btnSalvar.setToolTipText("Salvar alterações na planilha Excel");
        btnSalvar.addActionListener(e -> atualizarPlanilha());

        JButton btnValidar = new JButton("Validar [F5]");
        btnValidar.setToolTipText("Validar campos e posições");
        btnValidar.addActionListener(e -> validarCampos());

        JButton btnPrevisualizar = new JButton("Pré-visualizar XML [F6]");
        btnPrevisualizar.setToolTipText("Visualizar XML gerado antes de salvar");
        btnPrevisualizar.addActionListener(e -> previsualizarXml());

        JButton btnGerarXml = new JButton("Gerar XML");
        btnGerarXml.setToolTipText("Gerar e salvar arquivo XML");
        btnGerarXml.addActionListener(e -> gerarXml());

        toolBar.add(btnAbrir);
        toolBar.addSeparator();
        toolBar.add(btnSalvar);
        toolBar.addSeparator();
        toolBar.add(btnValidar);
        toolBar.add(btnPrevisualizar);
        toolBar.addSeparator();
        toolBar.add(btnGerarXml);

        return toolBar;
    }

    private JPanel criarPainelPrincipal() {
        JPanel panel = new JPanel(new BorderLayout());

        // Painel de boas-vindas (exibido antes de abrir planilha)
        JPanel bemVindo = criarPainelBemVindo();

        JTabbedPane abas = new JTabbedPane();
        abas.addTab("Campos Entrada", camposPanel);

        // Usa CardLayout para alternar entre bem-vindo e conteúdo
        JPanel cards = new JPanel(new CardLayout());
        cards.add(bemVindo, "bemVindo");
        cards.add(abas, "conteudo");

        // Guarda referência para trocar card quando planilha for aberta
        this.cardPanel = cards;
        this.cardLayout = (CardLayout) cards.getLayout();

        panel.add(cards, BorderLayout.CENTER);
        return panel;
    }

    private JPanel cardPanel;
    private CardLayout cardLayout;

    private JPanel criarPainelBemVindo() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0; gbc.gridy = 0;
        gbc.insets = new Insets(8, 8, 8, 8);

        JLabel titulo = new JLabel("AlteraEventos");
        titulo.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 28));
        titulo.setForeground(new Color(0x2C3E50));

        JLabel subtitulo = new JLabel("Gerador de XML a partir de Planilhas de Eventos");
        subtitulo.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 16));
        subtitulo.setForeground(Color.GRAY);

        JButton btnAbrirGrande = new JButton("Abrir Planilha (.xlsx)");
        btnAbrirGrande.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 14));
        btnAbrirGrande.setPreferredSize(new Dimension(250, 50));
        btnAbrirGrande.addActionListener(e -> abrirPlanilha());

        JLabel instrucoes = new JLabel("<html><center>"
                + "1. Abra uma planilha Excel (.xlsx) com as abas 'Campos Entrada'<br>"
                + "2. Configure os campos e preencha os valores para o XML<br>"
                + "3. Valide os campos e gere o XML final"
                + "</center></html>");
        instrucoes.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 13));
        instrucoes.setForeground(new Color(0x555555));

        panel.add(titulo, gbc); gbc.gridy++;
        panel.add(subtitulo, gbc); gbc.gridy++;
        gbc.insets = new Insets(24, 8, 8, 8);
        panel.add(btnAbrirGrande, gbc); gbc.gridy++;
        gbc.insets = new Insets(16, 8, 8, 8);
        panel.add(instrucoes, gbc);

        return panel;
    }

    private JPanel criarStatusBar() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(1, 0, 0, 0, Color.LIGHT_GRAY),
                BorderFactory.createEmptyBorder(2, 0, 2, 0)));

        panel.add(statusLabel, BorderLayout.WEST);
        panel.add(arquivoLabel, BorderLayout.EAST);
        return panel;
    }

    // =========================================================
    // Ações Principais
    // =========================================================

    private void abrirPlanilha() {
        if (alteracoesPendentes) {
            int opcao = JOptionPane.showConfirmDialog(this,
                    "Existem alterações não salvas. Deseja abrir outro arquivo mesmo assim?",
                    "Alterações pendentes", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
            if (opcao != JOptionPane.YES_OPTION) return;
        }

        JFileChooser chooser = new JFileChooser();
        String ultimoDir = prefs.get(PREF_ULTIMO_DIRETORIO, System.getProperty("user.home"));
        chooser.setCurrentDirectory(new File(ultimoDir));
        chooser.setDialogTitle("Abrir Planilha");
        chooser.setFileFilter(new FileNameExtensionFilter(
                "Planilhas Excel (*.xlsx)", "xlsx"));
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;

        File arquivo = chooser.getSelectedFile();
        prefs.put(PREF_ULTIMO_DIRETORIO, arquivo.getParent());

        carregarArquivo(arquivo);
    }

    private void carregarArquivo(File arquivo) {
        statusLabel.setText("Carregando planilha...");
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        SwingWorker<List<CampoEntrada>, Void> worker = new SwingWorker<>() {
            @Override
            protected List<CampoEntrada> doInBackground() throws Exception {
                return planilhaService.lerPlanilha(arquivo);
            }

            @Override
            protected void done() {
                setCursor(Cursor.getDefaultCursor());
                try {
                    List<CampoEntrada> campos = get();
                    arquivoAtual = arquivo;
                    camposPanel.carregarCampos(campos);
                    cardLayout.show(cardPanel, "conteudo");
                    alteracoesPendentes = false;

                    String msg = String.format("Planilha carregada: %d campos encontrados.", campos.size());
                    statusLabel.setText(msg);
                    arquivoLabel.setText(arquivo.getName());
                    setTitle(TITULO + " - " + arquivo.getName());

                } catch (Exception ex) {
                    String erro = extrairMensagemErro(ex);
                    statusLabel.setText("Erro ao carregar planilha.");
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Erro ao carregar a planilha:\n\n" + erro,
                            "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        };
        worker.execute();
    }

    private void atualizarPlanilha() {
        if (arquivoAtual == null) {
            JOptionPane.showMessageDialog(this, "Nenhuma planilha carregada.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        List<CampoEntrada> campos = camposPanel.getCampos();
        if (campos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Nenhum campo para salvar.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Validação antes de salvar
        ResultadoValidacao resultado = validacaoService.validar(campos);
        if (!resultado.isValido()) {
            int opcao = JOptionPane.showConfirmDialog(this,
                    "Existem erros de validação. Deseja salvar mesmo assim?\n\n"
                    + resultado.getRelatorio(),
                    "Validação com erros", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
            if (opcao != JOptionPane.YES_OPTION) return;
        }

        statusLabel.setText("Salvando planilha...");
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        SwingWorker<Void, Void> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground() throws Exception {
                planilhaService.salvarPlanilha(arquivoAtual, campos);
                return null;
            }

            @Override
            protected void done() {
                setCursor(Cursor.getDefaultCursor());
                try {
                    get();
                    alteracoesPendentes = false;
                    statusLabel.setText("Planilha salva com sucesso: " + arquivoAtual.getName());
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Planilha atualizada com sucesso!\n" + arquivoAtual.getAbsolutePath(),
                            "Planilha Salva", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception ex) {
                    statusLabel.setText("Erro ao salvar planilha.");
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Erro ao salvar a planilha:\n\n" + extrairMensagemErro(ex),
                            "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        };
        worker.execute();
    }

    private void validarCampos() {
        List<CampoEntrada> campos = camposPanel.getCampos();
        if (campos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Nenhuma planilha carregada.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        ResultadoValidacao resultado = validacaoService.validar(campos);
        statusLabel.setText("Validação: " + resultado.getStatusResumo());

        JTextArea textArea = new JTextArea(resultado.getRelatorio(), 22, 65);
        textArea.setEditable(false);
        textArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        JScrollPane scroll = new JScrollPane(textArea);
        scroll.setPreferredSize(new Dimension(700, 400));

        JOptionPane.showMessageDialog(this, scroll,
                resultado.isValido() ? "Validação Concluída" : "Validação com Problemas",
                resultado.isValido() ? JOptionPane.INFORMATION_MESSAGE : JOptionPane.WARNING_MESSAGE);
    }

    private void previsualizarXml() {
        List<CampoEntrada> campos = camposPanel.getCampos();
        if (campos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Nenhuma planilha carregada.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try {
            String xml = xmlService.gerarXmlString(campos);

            JTextArea textArea = new JTextArea(xml, 30, 80);
            textArea.setEditable(false);
            textArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
            textArea.setCaretPosition(0);

            JScrollPane scroll = new JScrollPane(textArea);
            scroll.setPreferredSize(new Dimension(800, 550));

            JOptionPane.showMessageDialog(this, scroll,
                    "Pré-visualização do XML", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "Erro ao gerar XML:\n" + extrairMensagemErro(ex),
                    "Erro", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void gerarXml() {
        List<CampoEntrada> campos = camposPanel.getCampos();
        if (campos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Nenhuma planilha carregada.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Valida antes de gerar
        ResultadoValidacao resultado = validacaoService.validar(campos);
        if (!resultado.isValido()) {
            int opcao = JOptionPane.showConfirmDialog(this,
                    "Existem erros de validação. Deseja gerar o XML mesmo assim?\n\n"
                    + resultado.getRelatorio(),
                    "Validação com erros", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
            if (opcao != JOptionPane.YES_OPTION) return;
        }

        // Escolhe local para salvar
        JFileChooser chooser = new JFileChooser();
        String ultimoDir = prefs.get(PREF_ULTIMO_DIRETORIO, System.getProperty("user.home"));
        chooser.setCurrentDirectory(new File(ultimoDir));
        chooser.setDialogTitle("Salvar XML");
        chooser.setFileFilter(new FileNameExtensionFilter("Arquivo XML (*.xml)", "xml"));
        chooser.setAcceptAllFileFilterUsed(false);

        // Sugere nome baseado no arquivo da planilha
        if (arquivoAtual != null) {
            String nomeSugerido = arquivoAtual.getName().replaceAll("\\.[^.]+$", "") + ".xml";
            chooser.setSelectedFile(new File(arquivoAtual.getParent(), nomeSugerido));
        } else {
            chooser.setSelectedFile(new File("evento.xml"));
        }

        if (chooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) return;

        File arquivoXml = chooser.getSelectedFile();
        if (!arquivoXml.getName().toLowerCase().endsWith(".xml")) {
            arquivoXml = new File(arquivoXml.getAbsolutePath() + ".xml");
        }

        final File arquivoFinal = arquivoXml;
        statusLabel.setText("Gerando XML...");
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        SwingWorker<Void, Void> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground() throws Exception {
                xmlService.salvarXml(campos, arquivoFinal);
                return null;
            }

            @Override
            protected void done() {
                setCursor(Cursor.getDefaultCursor());
                try {
                    get();
                    statusLabel.setText("XML gerado: " + arquivoFinal.getName());
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "XML gerado com sucesso!\n\n" + arquivoFinal.getAbsolutePath(),
                            "XML Gerado", JOptionPane.INFORMATION_MESSAGE);
                    prefs.put(PREF_ULTIMO_DIRETORIO, arquivoFinal.getParent());
                } catch (Exception ex) {
                    statusLabel.setText("Erro ao gerar XML.");
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Erro ao gerar XML:\n\n" + extrairMensagemErro(ex),
                            "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        };
        worker.execute();
    }

    private void mostrarSobre() {
        JOptionPane.showMessageDialog(this,
                "<html><center>"
                + "<b>AlteraEventos v1.0</b><br>"
                + "Gerador de XML a partir de Planilhas de Eventos<br><br>"
                + "Fluxo:<br>"
                + "1. Abrir planilha .xlsx<br>"
                + "2. Configurar campos na aba 'Campos Entrada'<br>"
                + "3. Validar e gerar XML<br><br>"
                + "Desenvolvido em Java com Apache POI"
                + "</center></html>",
                "Sobre", JOptionPane.INFORMATION_MESSAGE);
    }

    private void fecharAplicacao() {
        if (alteracoesPendentes) {
            int opcao = JOptionPane.showConfirmDialog(this,
                    "Existem alterações não salvas na planilha. Deseja sair mesmo assim?",
                    "Sair", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
            if (opcao != JOptionPane.YES_OPTION) return;
        }
        dispose();
        System.exit(0);
    }

    /**
     * Abre um arquivo diretamente (usado ao passar arquivo por argumento de linha de comando).
     */
    public void abrirArquivo(File arquivo) {
        carregarArquivo(arquivo);
    }

    private String extrairMensagemErro(Exception ex) {
        Throwable causa = ex.getCause();
        String msg = causa != null ? causa.getMessage() : ex.getMessage();
        return msg != null ? msg : ex.getClass().getSimpleName();
    }
}
