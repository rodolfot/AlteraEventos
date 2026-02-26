package com.alteraeventos.service;

import com.alteraeventos.model.CampoEntrada;
import com.alteraeventos.model.ResultadoValidacao;

import java.util.Comparator;
import java.util.List;

/**
 * Serviço de validação dos campos de entrada da planilha.
 * Verifica:
 *  - Consistência entre TamanhoCampo, PosicaoInicial e PosicaoFinal
 *  - Continuidade das posições (sem gaps ou sobreposições)
 *  - Campos obrigatórios sem valor
 */
public class ValidacaoService {

    /**
     * Valida a lista completa de campos.
     */
    public ResultadoValidacao validar(List<CampoEntrada> campos) {
        ResultadoValidacao resultado = new ResultadoValidacao();

        // Filtra apenas campos de entrada com posição definida
        List<CampoEntrada> camposEntrada = campos.stream()
                .filter(CampoEntrada::isEntradaComPosicao)
                .sorted(Comparator.comparing(CampoEntrada::getPosicaoInicial))
                .toList();

        if (camposEntrada.isEmpty()) {
            resultado.addAviso("Nenhum campo de entrada (Entrada=S) com posição definida encontrado.");
            return resultado;
        }

        resultado.setTotalCampos(camposEntrada.size());

        // Valida cada campo individualmente
        for (CampoEntrada campo : camposEntrada) {
            validarCampoIndividual(campo, resultado);
        }

        // Valida continuidade das posições
        validarContinuidade(camposEntrada, resultado);

        // Calcula tamanho total
        int totalSize = camposEntrada.stream()
                .mapToInt(c -> c.getTamanhoCampo() != null ? c.getTamanhoCampo() : 0)
                .sum();
        resultado.setTotalSize(totalSize);

        // Resumo final
        if (resultado.isValido()) {
            resultado.addInfo(String.format(
                    "Validação concluída: %d campos de entrada | Tamanho total do layout: %d bytes",
                    camposEntrada.size(), totalSize));
        }

        return resultado;
    }

    private void validarCampoIndividual(CampoEntrada campo, ResultadoValidacao resultado) {
        int posInicial = campo.getPosicaoInicial();
        int tamanho = campo.getTamanhoCampo();

        // Verifica tamanho positivo
        if (tamanho <= 0) {
            resultado.addErro(String.format(
                    "Linha %d | Campo '%s': TamanhoCampo inválido (%d). Deve ser maior que 0.",
                    campo.getLinha(), campo.getNomeCampo(), tamanho));
        }

        // Verifica posição inicial positiva
        if (posInicial <= 0) {
            resultado.addErro(String.format(
                    "Linha %d | Campo '%s': PosicaoInicial inválida (%d). Deve ser maior que 0.",
                    campo.getLinha(), campo.getNomeCampo(), posInicial));
        }

        // Verifica se PosicaoFinal confere com a fórmula
        if (campo.getPosicaoFinal() != null) {
            int posFinlEsperada = posInicial + tamanho - 1;
            if (campo.getPosicaoFinal() != posFinlEsperada) {
                resultado.addErro(String.format(
                        "Linha %d | Campo '%s': PosicaoFinal=%d mas esperado=%d (PosInicial=%d + Tamanho=%d - 1)",
                        campo.getLinha(), campo.getNomeCampo(),
                        campo.getPosicaoFinal(), posFinlEsperada, posInicial, tamanho));
            }
        }

        // Avisa sobre campos obrigatórios sem valor
        if ("S".equalsIgnoreCase(campo.getCampoObrigatorio())) {
            String valor = campo.getValorUsuario();
            if (valor == null || valor.isBlank()) {
                valor = campo.getValorPadrao();
            }
            if (valor == null || valor.isBlank()) {
                resultado.addAviso(String.format(
                        "Linha %d | Campo obrigatório '%s' sem valor preenchido.",
                        campo.getLinha(), campo.getNomeCampo()));
            }
        }

        // Valida tamanho do valor em relação ao campo
        String valor = campo.getValorUsuario();
        if (valor != null && !valor.isBlank() && valor.length() > tamanho) {
            resultado.addAviso(String.format(
                    "Linha %d | Campo '%s': valor (%d chars) excede o tamanho do campo (%d) e será truncado.",
                    campo.getLinha(), campo.getNomeCampo(), valor.length(), tamanho));
        }
    }

    private void validarContinuidade(List<CampoEntrada> campos, ResultadoValidacao resultado) {
        if (campos.isEmpty()) return;

        // Verifica se começa na posição 1
        int primeiraPosicao = campos.get(0).getPosicaoInicial();
        if (primeiraPosicao != 1) {
            resultado.addAviso(String.format(
                    "O layout não começa na posição 1. Primeiro campo '%s' começa na posição %d.",
                    campos.get(0).getNomeCampo(), primeiraPosicao));
        }

        // Verifica continuidade entre campos consecutivos
        for (int i = 1; i < campos.size(); i++) {
            CampoEntrada anterior = campos.get(i - 1);
            CampoEntrada atual = campos.get(i);

            int posicaoEsperada = anterior.getPosicaoInicial() + anterior.getTamanhoCampo();

            if (atual.getPosicaoInicial() > posicaoEsperada) {
                resultado.addAviso(String.format(
                        "Gap detectado: campo '%s' (termina em %d) → campo '%s' (começa em %d). Gap de %d byte(s).",
                        anterior.getNomeCampo(), anterior.getPosicaoInicial() + anterior.getTamanhoCampo() - 1,
                        atual.getNomeCampo(), atual.getPosicaoInicial(),
                        atual.getPosicaoInicial() - posicaoEsperada));
            } else if (atual.getPosicaoInicial() < posicaoEsperada) {
                resultado.addErro(String.format(
                        "Sobreposição detectada: campo '%s' (termina em %d) sobrepõe campo '%s' (começa em %d).",
                        anterior.getNomeCampo(), anterior.getPosicaoInicial() + anterior.getTamanhoCampo() - 1,
                        atual.getNomeCampo(), atual.getPosicaoInicial()));
            }
        }
    }

    /**
     * Recalcula as posições iniciais e finais de todos os campos em sequência.
     * Útil para corrigir automaticamente gaps ou sobreposições.
     */
    public void recalcularPosicoes(List<CampoEntrada> campos) {
        List<CampoEntrada> camposEntrada = campos.stream()
                .filter(CampoEntrada::isEntradaComPosicao)
                .sorted(Comparator.comparing(CampoEntrada::getPosicaoInicial))
                .toList();

        int posicaoAtual = 1;
        for (CampoEntrada campo : camposEntrada) {
            campo.setPosicaoInicial(posicaoAtual);
            campo.setPosicaoFinal(posicaoAtual + campo.getTamanhoCampo() - 1);
            posicaoAtual += campo.getTamanhoCampo();
        }
    }
}
