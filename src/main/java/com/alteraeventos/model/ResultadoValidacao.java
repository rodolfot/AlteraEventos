package com.alteraeventos.model;

import java.util.ArrayList;
import java.util.List;

/**
 * Resultado de uma validação dos campos da planilha.
 */
public class ResultadoValidacao {

    private final List<String> erros = new ArrayList<>();
    private final List<String> avisos = new ArrayList<>();
    private final List<String> infos = new ArrayList<>();
    private int totalSize;
    private int totalCampos;

    public void addErro(String msg) { erros.add(msg); }
    public void addAviso(String msg) { avisos.add(msg); }
    public void addInfo(String msg) { infos.add(msg); }

    public boolean isValido() { return erros.isEmpty(); }
    public boolean temAvisos() { return !avisos.isEmpty(); }

    public List<String> getErros() { return erros; }
    public List<String> getAvisos() { return avisos; }
    public List<String> getInfos() { return infos; }

    public int getTotalSize() { return totalSize; }
    public void setTotalSize(int totalSize) { this.totalSize = totalSize; }

    public int getTotalCampos() { return totalCampos; }
    public void setTotalCampos(int totalCampos) { this.totalCampos = totalCampos; }

    /**
     * Gera um relatório textual completo da validação.
     */
    public String getRelatorio() {
        StringBuilder sb = new StringBuilder();

        if (!infos.isEmpty()) {
            sb.append("=== INFORMAÇÕES ===\n");
            infos.forEach(i -> sb.append("  ✔ ").append(i).append("\n"));
        }

        if (!avisos.isEmpty()) {
            sb.append("\n=== AVISOS ===\n");
            avisos.forEach(a -> sb.append("  ⚠ ").append(a).append("\n"));
        }

        if (!erros.isEmpty()) {
            sb.append("\n=== ERROS ===\n");
            erros.forEach(e -> sb.append("  ✗ ").append(e).append("\n"));
        }

        if (erros.isEmpty() && avisos.isEmpty()) {
            sb.append("✔ Validação concluída sem problemas.\n");
        }

        if (totalCampos > 0) {
            sb.append(String.format("\nTotal: %d campos | Tamanho do layout: %d bytes\n",
                    totalCampos, totalSize));
        }

        return sb.toString();
    }

    /**
     * Retorna apenas a linha de status resumida.
     */
    public String getStatusResumo() {
        if (!isValido()) {
            return String.format("INVÁLIDO: %d erro(s), %d aviso(s)", erros.size(), avisos.size());
        } else if (temAvisos()) {
            return String.format("OK com avisos: %d aviso(s) | %d campos | %d bytes",
                    avisos.size(), totalCampos, totalSize);
        } else {
            return String.format("OK: %d campos | %d bytes", totalCampos, totalSize);
        }
    }
}
