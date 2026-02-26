package com.alteraeventos.service;

import com.alteraeventos.model.CampoEntrada;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.StringWriter;
import java.util.Comparator;
import java.util.List;

/**
 * Serviço de geração de XML a partir dos campos de entrada.
 *
 * Estrutura do XML gerado:
 * <evento tamanhoTotal="N">
 *   <campos>
 *     <NomeCampo id="1" tipo="TEXTO" tamanho="8" posicaoInicial="1" posicaoFinal="8"
 *                alinhamento="BRANCO_DIREITA" obrigatorio="S"
 *                descricao="...">VALOR_FORMATADO</NomeCampo>
 *     ...
 *   </campos>
 * </evento>
 */
public class XmlService {

    /**
     * Gera o XML como String.
     */
    public String gerarXmlString(List<CampoEntrada> campos) throws Exception {
        Document doc = construirDocumento(campos);
        return documentoParaString(doc);
    }

    /**
     * Gera o XML e salva em arquivo.
     */
    public void salvarXml(List<CampoEntrada> campos, File arquivo) throws Exception {
        Document doc = construirDocumento(campos);
        salvarDocumento(doc, arquivo);
    }

    private Document construirDocumento(List<CampoEntrada> campos) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.newDocument();

        // Filtra e ordena campos de entrada
        List<CampoEntrada> camposEntrada = campos.stream()
                .filter(CampoEntrada::isEntradaComPosicao)
                .sorted(Comparator.comparing(CampoEntrada::getPosicaoInicial))
                .toList();

        // Elemento raiz
        Element root = doc.createElement("evento");
        doc.appendChild(root);

        int totalSize = camposEntrada.stream()
                .mapToInt(c -> c.getTamanhoCampo() != null ? c.getTamanhoCampo() : 0)
                .sum();
        root.setAttribute("tamanhoTotal", String.valueOf(totalSize));
        root.setAttribute("totalCampos", String.valueOf(camposEntrada.size()));

        // Container de campos
        Element camposEl = doc.createElement("campos");
        root.appendChild(camposEl);

        for (CampoEntrada campo : camposEntrada) {
            Element campoEl = doc.createElement(sanitizarNomeXml(campo.getNomeCampo()));

            // Atributos posicionais e descritivos
            if (campo.getIdentificadorCampo() != null) {
                campoEl.setAttribute("id", String.valueOf(campo.getIdentificadorCampo()));
            }
            campoEl.setAttribute("tipo", campo.getTipoCampo());
            if (campo.getTamanhoCampo() != null) {
                campoEl.setAttribute("tamanho", String.valueOf(campo.getTamanhoCampo()));
            }
            if (campo.getPosicaoInicial() != null) {
                campoEl.setAttribute("posicaoInicial", String.valueOf(campo.getPosicaoInicial()));
            }
            if (campo.getPosicaoFinal() != null) {
                campoEl.setAttribute("posicaoFinal", String.valueOf(campo.getPosicaoFinal()));
            } else if (campo.getPosicaoInicial() != null && campo.getTamanhoCampo() != null) {
                campoEl.setAttribute("posicaoFinal",
                        String.valueOf(campo.getPosicaoInicial() + campo.getTamanhoCampo() - 1));
            }
            if (!campo.getAlinhamentoCampo().isBlank()) {
                campoEl.setAttribute("alinhamento", campo.getAlinhamentoCampo());
            }
            if (!campo.getCampoObrigatorio().isBlank()) {
                campoEl.setAttribute("obrigatorio", campo.getCampoObrigatorio());
            }
            if (!campo.getDescricaoCampo().isBlank()) {
                campoEl.setAttribute("descricao", campo.getDescricaoCampo());
            }
            if (!campo.getNomeColuna().isBlank()) {
                campoEl.setAttribute("colunaDB", campo.getNomeColuna());
            }

            // Valor do campo: usa valorUsuario, senão valorPadrao, senão vazio
            String valor = campo.getValorUsuario();
            if (valor == null || valor.isBlank()) {
                valor = campo.getValorPadrao();
                if (valor == null || valor.isBlank()) {
                    valor = campo.getDefaultValue();
                }
            }

            // Aplica alinhamento/padding conforme configuração do campo
            if (campo.getTamanhoCampo() != null) {
                valor = aplicarAlinhamento(valor != null ? valor : "", campo.getTamanhoCampo(),
                        campo.getAlinhamentoCampo(), campo.getTipoCampo());
            }

            campoEl.setTextContent(valor != null ? valor : "");
            camposEl.appendChild(campoEl);
        }

        return doc;
    }

    /**
     * Aplica o alinhamento/padding conforme a regra da planilha.
     *
     * BRANCO_DIREITA  → valor alinhado à direita com espaços à esquerda
     * BRANCO_ESQUERDA → valor alinhado à esquerda com espaços à direita
     * ZERO_ESQUERDA   → valor alinhado à direita com zeros à esquerda
     * ZERO_DIREITA    → valor alinhado à esquerda com zeros à direita
     */
    public String aplicarAlinhamento(String valor, int tamanho, String alinhamento, String tipoCampo) {
        if (valor == null) valor = "";

        // Trunca se exceder
        if (valor.length() > tamanho) {
            return valor.substring(0, tamanho);
        }

        int diff = tamanho - valor.length();
        if (diff == 0) return valor;

        String padChar;
        boolean padEsquerda;

        String alin = alinhamento != null ? alinhamento.toUpperCase().trim() : "";

        // Se não houver alinhamento definido, usa padrão por tipo
        if (alin.isBlank()) {
            boolean numerico = tipoCampo != null &&
                    (tipoCampo.contains("INTEIRO") || tipoCampo.contains("DECIMAL") || tipoCampo.contains("NUMERO"));
            alin = numerico ? "ZERO_ESQUERDA" : "BRANCO_ESQUERDA";
        }

        return switch (alin) {
            case "BRANCO_DIREITA" -> " ".repeat(diff) + valor;
            case "BRANCO_ESQUERDA" -> valor + " ".repeat(diff);
            case "ZERO_ESQUERDA" -> "0".repeat(diff) + valor;
            case "ZERO_DIREITA" -> valor + "0".repeat(diff);
            default -> valor + " ".repeat(diff);
        };
    }

    /**
     * Sanitiza o nome do campo para ser um nome XML válido.
     * Remove caracteres inválidos e garante que começa com letra ou underscore.
     */
    private String sanitizarNomeXml(String nome) {
        if (nome == null || nome.isBlank()) return "campo";

        // Substitui caracteres inválidos por underscore
        String sanitizado = nome.trim()
                .replaceAll("[^a-zA-Z0-9_\\-.]", "_")
                .replaceAll("^[^a-zA-Z_]", "_$0");

        return sanitizado.isBlank() ? "campo" : sanitizado;
    }

    private String documentoParaString(Document doc) throws Exception {
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer transformer = tf.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

        StringWriter writer = new StringWriter();
        transformer.transform(new DOMSource(doc), new StreamResult(writer));
        return writer.toString();
    }

    private void salvarDocumento(Document doc, File arquivo) throws Exception {
        TransformerFactory tf = TransformerFactory.newInstance();
        Transformer transformer = tf.newTransformer();
        transformer.setOutputProperty(OutputKeys.INDENT, "yes");
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");

        transformer.transform(new DOMSource(doc), new StreamResult(arquivo));
    }
}
