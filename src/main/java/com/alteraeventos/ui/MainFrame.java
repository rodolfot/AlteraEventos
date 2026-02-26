package com.alteraeventos.ui;

import com.alteraeventos.model.CampoEntrada;
import com.alteraeventos.model.ResultadoValidacao;
import com.alteraeventos.service.PlanilhaService;
import com.alteraeventos.service.ValidacaoService;
import com.alteraeventos.service.XmlService;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.prefs.Preferences;

/**
 * Janela principal do sistema AlteraEventos.
 * Compatível com Java 8.
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

    private JPanel cardPanel;
    private CardLayout cardLayout;

    private File arquivoAtual;

    public MainFrame() {
        setTitle(TITULO);
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);
        setMinimumSize(new Dimension(1100, 700));
        setPreferredSize(new Dimension(1300, 800));

        camposPanel = new CamposEntradaPanel();
        camposPanel.setOnAtualizarPlanilha(new Runnable() {
            public void run() { atualizarPlanilha(); }
        });
        camposPanel.setOnGerarXml(new Runnable() {
            public void run() { gerarXml(); }
        });

        statusLabel = new JLabel("Pronto. Abra uma planilha para começar.");
        statusLabel.setBorder(BorderFactory.createEmptyBorder(2, 8, 2, 8));

        arquivoLabel = new JLabel("Nenhum arquivo carregado");
        arquivoLabel.setForeground(Color.GRAY);
        arquivoLabel.setBorder(BorderFactory.createEmptyBorder(2, 8, 2, 8));

        setJMenuBar(criarMenuBar());
        add(criarToolBar(), BorderLayout.NORTH);
        add(criarPainelPrincipal(), BorderLayout.CENTER);
        add(criarStatusBar(), BorderLayout.SOUTH);

        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) { fecharAplicacao(); }
        });

        pack();
        setLocationRelativeTo(null);
    }

    private JMenuBar criarMenuBar() {
        JMenuBar menuBar = new JMenuBar();

        JMenu menuArquivo = new JMenu("Arquivo");
        menuArquivo.setMnemonic('A');

        JMenuItem itemAbrir = new JMenuItem("Abrir Planilha...");
        itemAbrir.setAccelerator(KeyStroke.getKeyStroke("control O"));
        itemAbrir.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { abrirPlanilha(); }
        });

        JMenuItem itemSalvar = new JMenuItem("Atualizar Planilha");
        itemSalvar.setAccelerator(KeyStroke.getKeyStroke("control S"));
        itemSalvar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { atualizarPlanilha(); }
        });

        JMenuItem itemGerarXml = new JMenuItem("Gerar XML...");
        itemGerarXml.setAccelerator(KeyStroke.getKeyStroke("control G"));
        itemGerarXml.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { gerarXml(); }
        });

        JMenuItem itemSair = new JMenuItem("Sair");
        itemSair.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { fecharAplicacao(); }
        });

        menuArquivo.add(itemAbrir);
        menuArquivo.add(itemSalvar);
        menuArquivo.addSeparator();
        menuArquivo.add(itemGerarXml);
        menuArquivo.addSeparator();
        menuArquivo.add(itemSair);

        JMenu menuFerramentas = new JMenu("Ferramentas");
        menuFerramentas.setMnemonic('F');

        JMenuItem itemValidar = new JMenuItem("Validar Campos [F5]");
        itemValidar.setAccelerator(KeyStroke.getKeyStroke("F5"));
        itemValidar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { validarCampos(); }
        });

        JMenuItem itemPrevisualizar = new JMenuItem("Pré-visualizar XML [F6]");
        itemPrevisualizar.setAccelerator(KeyStroke.getKeyStroke("F6"));
        itemPrevisualizar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { previsualizarXml(); }
        });

        menuFerramentas.add(itemValidar);
        menuFerramentas.add(itemPrevisualizar);

        JMenu menuAjuda = new JMenu("Ajuda");
        JMenuItem itemSobre = new JMenuItem("Sobre");
        itemSobre.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { mostrarSobre(); }
        });
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
        btnAbrir.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { abrirPlanilha(); }
        });

        JButton btnSalvar = new JButton("Atualizar Planilha");
        btnSalvar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { atualizarPlanilha(); }
        });

        JButton btnValidar = new JButton("Validar [F5]");
        btnValidar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { validarCampos(); }
        });

        JButton btnPrevisualizar = new JButton("Pré-visualizar XML [F6]");
        btnPrevisualizar.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { previsualizarXml(); }
        });

        JButton btnGerarXml = new JButton("Gerar XML");
        btnGerarXml.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { gerarXml(); }
        });

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

        JTabbedPane abas = new JTabbedPane();
        abas.addTab("Campos Entrada", camposPanel);

        cardPanel = new JPanel(new CardLayout());
        cardLayout = (CardLayout) cardPanel.getLayout();
        cardPanel.add(criarPainelBemVindo(), "bemVindo");
        cardPanel.add(abas, "conteudo");

        panel.add(cardPanel, BorderLayout.CENTER);
        return panel;
    }

    private JPanel criarPainelBemVindo() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0; gbc.insets = new Insets(8, 8, 8, 8);

        JLabel titulo = new JLabel("AlteraEventos");
        titulo.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 28));
        titulo.setForeground(new Color(0x2C3E50));

        JLabel subtitulo = new JLabel("Gerador de XML a partir de Planilhas de Eventos");
        subtitulo.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 16));
        subtitulo.setForeground(Color.GRAY);

        JButton btnAbrir = new JButton("Abrir Planilha (.xlsx)");
        btnAbrir.setFont(new Font(Font.SANS_SERIF, Font.BOLD, 14));
        btnAbrir.setPreferredSize(new Dimension(250, 50));
        btnAbrir.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) { abrirPlanilha(); }
        });

        JLabel instrucoes = new JLabel("<html><center>"
                + "1. Abra uma planilha Excel (.xlsx) com a aba 'Campos Entrada'<br>"
                + "2. Configure os campos e preencha os valores<br>"
                + "3. Valide e gere o XML final"
                + "</center></html>");
        instrucoes.setFont(new Font(Font.SANS_SERIF, Font.PLAIN, 13));

        gbc.gridy = 0; panel.add(titulo, gbc);
        gbc.gridy = 1; panel.add(subtitulo, gbc);
        gbc.gridy = 2; gbc.insets = new Insets(24, 8, 8, 8); panel.add(btnAbrir, gbc);
        gbc.gridy = 3; gbc.insets = new Insets(16, 8, 8, 8); panel.add(instrucoes, gbc);
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
    // Ações
    // =========================================================

    private void abrirPlanilha() {
        JFileChooser chooser = new JFileChooser();
        String ultimoDir = prefs.get(PREF_ULTIMO_DIRETORIO, System.getProperty("user.home"));
        chooser.setCurrentDirectory(new File(ultimoDir));
        chooser.setDialogTitle("Abrir Planilha");
        chooser.setFileFilter(new FileNameExtensionFilter("Planilhas Excel (*.xlsx)", "xlsx"));
        chooser.setAcceptAllFileFilterUsed(false);

        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) return;

        File arquivo = chooser.getSelectedFile();
        prefs.put(PREF_ULTIMO_DIRETORIO, arquivo.getParent());
        carregarArquivo(arquivo);
    }

    void carregarArquivo(final File arquivo) {
        statusLabel.setText("Carregando planilha...");
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        new SwingWorker<List<CampoEntrada>, Void>() {
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

                    String msg = String.format("Planilha carregada: %d campos encontrados.", campos.size());
                    statusLabel.setText(msg);
                    arquivoLabel.setText(arquivo.getName());
                    setTitle(TITULO + " - " + arquivo.getName());
                } catch (InterruptedException | ExecutionException ex) {
                    statusLabel.setText("Erro ao carregar planilha.");
                    Throwable causa = ex.getCause() != null ? ex.getCause() : ex;
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Erro ao carregar a planilha:\n\n" + causa.getMessage(),
                            "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
    }

    private void atualizarPlanilha() {
        if (arquivoAtual == null) {
            JOptionPane.showMessageDialog(this, "Nenhuma planilha carregada.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        final List<CampoEntrada> campos = camposPanel.getCampos();
        if (campos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Nenhum campo para salvar.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        ResultadoValidacao resultado = validacaoService.validar(campos);
        if (!resultado.isValido()) {
            int opcao = JOptionPane.showConfirmDialog(this,
                    "Existem erros de validação. Deseja salvar mesmo assim?\n\n" + resultado.getRelatorio(),
                    "Validação com erros", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
            if (opcao != JOptionPane.YES_OPTION) return;
        }

        statusLabel.setText("Salvando planilha...");
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        new SwingWorker<Void, Void>() {
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
                    statusLabel.setText("Planilha salva: " + arquivoAtual.getName());
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Planilha atualizada com sucesso!\n" + arquivoAtual.getAbsolutePath(),
                            "Planilha Salva", JOptionPane.INFORMATION_MESSAGE);
                } catch (InterruptedException | ExecutionException ex) {
                    statusLabel.setText("Erro ao salvar planilha.");
                    Throwable causa = ex.getCause() != null ? ex.getCause() : ex;
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Erro ao salvar:\n\n" + causa.getMessage(),
                            "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
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
                    "Erro ao gerar XML:\n" + ex.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void gerarXml() {
        final List<CampoEntrada> campos = camposPanel.getCampos();
        if (campos.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Nenhuma planilha carregada.",
                    "Aviso", JOptionPane.WARNING_MESSAGE);
            return;
        }

        ResultadoValidacao resultado = validacaoService.validar(campos);
        if (!resultado.isValido()) {
            int opcao = JOptionPane.showConfirmDialog(this,
                    "Existem erros de validação. Deseja gerar o XML mesmo assim?\n\n" + resultado.getRelatorio(),
                    "Validação com erros", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
            if (opcao != JOptionPane.YES_OPTION) return;
        }

        JFileChooser chooser = new JFileChooser();
        String ultimoDir = prefs.get(PREF_ULTIMO_DIRETORIO, System.getProperty("user.home"));
        chooser.setCurrentDirectory(new File(ultimoDir));
        chooser.setDialogTitle("Salvar XML");
        chooser.setFileFilter(new FileNameExtensionFilter("Arquivo XML (*.xml)", "xml"));
        chooser.setAcceptAllFileFilterUsed(false);

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

        new SwingWorker<Void, Void>() {
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
                    prefs.put(PREF_ULTIMO_DIRETORIO, arquivoFinal.getParent());
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "XML gerado com sucesso!\n\n" + arquivoFinal.getAbsolutePath(),
                            "XML Gerado", JOptionPane.INFORMATION_MESSAGE);
                } catch (InterruptedException | ExecutionException ex) {
                    statusLabel.setText("Erro ao gerar XML.");
                    Throwable causa = ex.getCause() != null ? ex.getCause() : ex;
                    JOptionPane.showMessageDialog(MainFrame.this,
                            "Erro ao gerar XML:\n\n" + causa.getMessage(),
                            "Erro", JOptionPane.ERROR_MESSAGE);
                }
            }
        }.execute();
    }

    private void mostrarSobre() {
        JOptionPane.showMessageDialog(this,
                "<html><center>"
                + "<b>AlteraEventos v1.0</b><br>"
                + "Gerador de XML a partir de Planilhas de Eventos<br><br>"
                + "Compatível com Java 8+<br>"
                + "Desenvolvido com Apache POI"
                + "</center></html>",
                "Sobre", JOptionPane.INFORMATION_MESSAGE);
    }

    private void fecharAplicacao() {
        dispose();
        System.exit(0);
    }

    /**
     * Abre um arquivo diretamente (usado ao passar arquivo por argumento de linha de comando).
     */
    public void abrirArquivo(File arquivo) {
        carregarArquivo(arquivo);
    }
}
