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
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * Serviço de geração de XML a partir dos campos de entrada.
 * Compatível com Java 8.
 */
public class XmlService {

    public String gerarXmlString(List<CampoEntrada> campos) throws Exception {
        Document doc = construirDocumento(campos);
        return documentoParaString(doc);
    }

    public void salvarXml(List<CampoEntrada> campos, File arquivo) throws Exception {
        Document doc = construirDocumento(campos);
        salvarDocumento(doc, arquivo);
    }

    private Document construirDocumento(List<CampoEntrada> campos) throws Exception {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.newDocument();

        List<CampoEntrada> camposEntrada = filtrarEOrdenar(campos);

        Element root = doc.createElement("evento");
        doc.appendChild(root);

        int totalSize = 0;
        for (CampoEntrada c : camposEntrada) {
            totalSize += (c.getTamanhoCampo() != null ? c.getTamanhoCampo() : 0);
        }
        root.setAttribute("tamanhoTotal", String.valueOf(totalSize));
        root.setAttribute("totalCampos", String.valueOf(camposEntrada.size()));

        Element camposEl = doc.createElement("campos");
        root.appendChild(camposEl);

        for (CampoEntrada campo : camposEntrada) {
            Element campoEl = doc.createElement(sanitizarNomeXml(campo.getNomeCampo()));

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
            int posFinal = campo.getPosicaoFinal() != null
                    ? campo.getPosicaoFinal()
                    : campo.calcularPosicaoFinalEsperada();
            campoEl.setAttribute("posicaoFinal", String.valueOf(posFinal));

            if (!campo.getAlinhamentoCampo().trim().isEmpty()) {
                campoEl.setAttribute("alinhamento", campo.getAlinhamentoCampo());
            }
            if (!campo.getCampoObrigatorio().trim().isEmpty()) {
                campoEl.setAttribute("obrigatorio", campo.getCampoObrigatorio());
            }
            if (!campo.getDescricaoCampo().trim().isEmpty()) {
                campoEl.setAttribute("descricao", campo.getDescricaoCampo());
            }
            if (!campo.getNomeColuna().trim().isEmpty()) {
                campoEl.setAttribute("colunaDB", campo.getNomeColuna());
            }

            String valor = campo.getValorUsuario();
            if (valor == null || valor.trim().isEmpty()) {
                valor = campo.getValorPadrao();
                if (valor == null || valor.trim().isEmpty()) {
                    valor = campo.getDefaultValue();
                }
            }

            if (campo.getTamanhoCampo() != null) {
                valor = aplicarAlinhamento(
                        valor != null ? valor : "",
                        campo.getTamanhoCampo(),
                        campo.getAlinhamentoCampo(),
                        campo.getTipoCampo());
            }

            campoEl.setTextContent(valor != null ? valor : "");
            camposEl.appendChild(campoEl);
        }

        return doc;
    }

    /**
     * Aplica padding/alinhamento conforme a regra da planilha.
     */
    public String aplicarAlinhamento(String valor, int tamanho, String alinhamento, String tipoCampo) {
        if (valor == null) valor = "";
        if (valor.length() > tamanho) {
            return valor.substring(0, tamanho);
        }

        int diff = tamanho - valor.length();
        if (diff == 0) return valor;

        String alin = alinhamento != null ? alinhamento.toUpperCase().trim() : "";

        if (alin.isEmpty()) {
            boolean numerico = tipoCampo != null &&
                    (tipoCampo.contains("INTEIRO") || tipoCampo.contains("DECIMAL"));
            alin = numerico ? "ZERO_ESQUERDA" : "BRANCO_ESQUERDA";
        }

        switch (alin) {
            case "BRANCO_DIREITA":  return repeatChar(' ', diff) + valor;
            case "BRANCO_ESQUERDA": return valor + repeatChar(' ', diff);
            case "ZERO_ESQUERDA":   return repeatChar('0', diff) + valor;
            case "ZERO_DIREITA":    return valor + repeatChar('0', diff);
            default:                return valor + repeatChar(' ', diff);
        }
    }

    private String repeatChar(char c, int times) {
        if (times <= 0) return "";
        char[] chars = new char[times];
        java.util.Arrays.fill(chars, c);
        return new String(chars);
    }

    private String sanitizarNomeXml(String nome) {
        if (nome == null || nome.trim().isEmpty()) return "campo";
        String s = nome.trim().replaceAll("[^a-zA-Z0-9_\\-.]", "_");
        if (!s.isEmpty() && !Character.isLetter(s.charAt(0)) && s.charAt(0) != '_') {
            s = "_" + s;
        }
        return s.isEmpty() ? "campo" : s;
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
