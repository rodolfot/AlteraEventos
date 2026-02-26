package com.alteraeventos.service;

import com.alteraeventos.model.CampoEntrada;
import com.alteraeventos.model.ResultadoValidacao;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * Serviço de validação dos campos de entrada da planilha.
 * Compatível com Java 8.
 */
public class ValidacaoService {

    public ResultadoValidacao validar(List<CampoEntrada> campos) {
        ResultadoValidacao resultado = new ResultadoValidacao();

        List<CampoEntrada> camposEntrada = filtrarEOrdenar(campos);

        if (camposEntrada.isEmpty()) {
            resultado.addAviso("Nenhum campo de entrada (Entrada=S) com posição definida encontrado.");
            return resultado;
        }

        resultado.setTotalCampos(camposEntrada.size());

        for (CampoEntrada campo : camposEntrada) {
            validarCampoIndividual(campo, resultado);
        }

        validarContinuidade(camposEntrada, resultado);

        int totalSize = 0;
        for (CampoEntrada c : camposEntrada) {
            totalSize += (c.getTamanhoCampo() != null ? c.getTamanhoCampo() : 0);
        }
        resultado.setTotalSize(totalSize);

        if (resultado.isValido()) {
            resultado.addInfo(String.format(
                    "Validação concluída: %d campos | Tamanho total do layout: %d bytes",
                    camposEntrada.size(), totalSize));
        }

        return resultado;
    }

    private void validarCampoIndividual(CampoEntrada campo, ResultadoValidacao resultado) {
        int posInicial = campo.getPosicaoInicial();
        int tamanho = campo.getTamanhoCampo();

        if (tamanho <= 0) {
            resultado.addErro(String.format(
                    "Linha %d | Campo '%s': TamanhoCampo inválido (%d).",
                    campo.getLinha(), campo.getNomeCampo(), tamanho));
        }

        if (posInicial <= 0) {
            resultado.addErro(String.format(
                    "Linha %d | Campo '%s': PosicaoInicial inválida (%d).",
                    campo.getLinha(), campo.getNomeCampo(), posInicial));
        }

        if (campo.getPosicaoFinal() != null) {
            int posFinlEsperada = posInicial + tamanho - 1;
            if (campo.getPosicaoFinal() != posFinlEsperada) {
                resultado.addErro(String.format(
                        "Linha %d | Campo '%s': PosicaoFinal=%d mas esperado=%d (PosInicial=%d + Tamanho=%d - 1)",
                        campo.getLinha(), campo.getNomeCampo(),
                        campo.getPosicaoFinal(), posFinlEsperada, posInicial, tamanho));
            }
        }

        if ("S".equalsIgnoreCase(campo.getCampoObrigatorio())) {
            String valor = campo.getValorUsuario();
            if (valor == null || valor.trim().isEmpty()) {
                valor = campo.getValorPadrao();
            }
            if (valor == null || valor.trim().isEmpty()) {
                resultado.addAviso(String.format(
                        "Linha %d | Campo obrigatório '%s' sem valor preenchido.",
                        campo.getLinha(), campo.getNomeCampo()));
            }
        }

        String valor = campo.getValorUsuario();
        if (valor != null && !valor.trim().isEmpty() && valor.length() > tamanho) {
            resultado.addAviso(String.format(
                    "Linha %d | Campo '%s': valor (%d chars) excede tamanho (%d) e será truncado.",
                    campo.getLinha(), campo.getNomeCampo(), valor.length(), tamanho));
        }
    }

    private void validarContinuidade(List<CampoEntrada> campos, ResultadoValidacao resultado) {
        if (campos.isEmpty()) return;

        int primeiraPosicao = campos.get(0).getPosicaoInicial();
        if (primeiraPosicao != 1) {
            resultado.addAviso(String.format(
                    "O layout não começa na posição 1. Primeiro campo '%s' começa na posição %d.",
                    campos.get(0).getNomeCampo(), primeiraPosicao));
        }

        for (int i = 1; i < campos.size(); i++) {
            CampoEntrada anterior = campos.get(i - 1);
            CampoEntrada atual = campos.get(i);
            int posicaoEsperada = anterior.getPosicaoInicial() + anterior.getTamanhoCampo();

            if (atual.getPosicaoInicial() > posicaoEsperada) {
                resultado.addAviso(String.format(
                        "Gap detectado antes de '%s': esperado pos=%d, encontrado pos=%d. Gap de %d byte(s).",
                        atual.getNomeCampo(), posicaoEsperada, atual.getPosicaoInicial(),
                        atual.getPosicaoInicial() - posicaoEsperada));
            } else if (atual.getPosicaoInicial() < posicaoEsperada) {
                resultado.addErro(String.format(
                        "Sobreposição detectada: '%s' (termina em %d) sobrepõe '%s' (começa em %d).",
                        anterior.getNomeCampo(), anterior.getPosicaoInicial() + anterior.getTamanhoCampo() - 1,
                        atual.getNomeCampo(), atual.getPosicaoInicial()));
            }
        }
    }

    public void recalcularPosicoes(List<CampoEntrada> campos) {
        List<CampoEntrada> camposEntrada = filtrarEOrdenar(campos);
        int posicaoAtual = 1;
        for (CampoEntrada campo : camposEntrada) {
            campo.setPosicaoInicial(posicaoAtual);
            campo.setPosicaoFinal(posicaoAtual + campo.getTamanhoCampo() - 1);
            posicaoAtual += campo.getTamanhoCampo();
        }
    }

    private List<CampoEntrada> filtrarEOrdenar(List<CampoEntrada> campos) {
        List<CampoEntrada> resultado = new ArrayList<>();
        for (CampoEntrada c : campos) {
            if (c.isEntradaComPosicao()) {
                resultado.add(c);
            }
        }
        Collections.sort(resultado, new Comparator<CampoEntrada>() {
            @Override
            public int compare(CampoEntrada a, CampoEntrada b) {
                return Integer.compare(a.getPosicaoInicial(), b.getPosicaoInicial());
            }
        });
        return resultado;
    }
}
