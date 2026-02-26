package com.alteraeventos;

import com.alteraeventos.ui.MainFrame;

import javax.swing.*;

/**
 * Ponto de entrada do sistema AlteraEventos.
 *
 * Uso:
 *   java -jar AlteraEventos.jar
 *   java -jar AlteraEventos.jar caminho/para/planilha.xlsx
 */
public class Main {

    public static void main(String[] args) {
        // Habilita renderização de fontes para melhor legibilidade no Windows
        System.setProperty("awt.useSystemAAFontSettings", "on");
        System.setProperty("swing.aatext", "true");

        SwingUtilities.invokeLater(() -> {
            // Tenta usar o look and feel nativo do sistema operacional
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                // Mantém o look and feel padrão em caso de falha
            }

            MainFrame frame = new MainFrame();
            frame.setVisible(true);

            // Se um arquivo foi passado como argumento, abre automaticamente
            if (args.length > 0) {
                java.io.File arquivo = new java.io.File(args[0]);
                if (arquivo.exists() && arquivo.getName().endsWith(".xlsx")) {
                    // Dispara carregamento após a janela estar visível
                    SwingUtilities.invokeLater(() -> frame.abrirArquivo(arquivo));
                }
            }
        });
    }
}
