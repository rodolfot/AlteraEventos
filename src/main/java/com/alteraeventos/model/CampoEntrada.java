package com.alteraeventos.model;

/**
 * Representa um campo da aba "Campos Entrada" da planilha.
 * Cada instância corresponde a uma linha da planilha com as configurações do campo.
 */
public class CampoEntrada {

    // Número da linha na planilha (1-based)
    private int linha;

    // Flags de presença (S/N)
    private String entrada = "";
    private String persistencia = "";
    private String enriquecimento = "";
    private String mapaAtributo = "";
    private String saida = "";
    private String campoConcatenado = "";

    // Identificação e descrição
    private Integer identificadorCampo;
    private String nomeCampo = "";
    private String descricaoCampo = "";

    // Tipo e posicionamento
    private String tipoCampo = "";
    private Integer tamanhoCampo;
    private Integer posicaoInicial;
    private Integer posicaoFinal;

    // Configuração
    private String valorPadrao = "";
    private String alinhamentoCampo = "";
    private String campoObrigatorio = "";
    private String dominioCampo = "";
    private String mascaraCampo = "";

    // Banco de dados
    private String nomeTabela = "";
    private String nomeColuna = "";
    private String oracleDataType = "";
    private Integer dataLength;
    private Integer numberPrecision;
    private Integer numberScale;
    private String nullable = "";
    private String encrypted = "";
    private String unique = "";

    // Rule/Score attributes
    private String ruleAttribute = "";
    private String defaultValue = "";
    private String description = "";
    private String origin = "";
    private String eventAttribute = "";
    private String type = "";
    private String modelAttribute = "";
    private String scoreModelIn = "";

    // Valor preenchido pelo usuário para geração do XML
    private String valorUsuario = "";

    // =========================================================
    // Getters e Setters
    // =========================================================

    public int getLinha() { return linha; }
    public void setLinha(int linha) { this.linha = linha; }

    public String getEntrada() { return entrada; }
    public void setEntrada(String v) { this.entrada = v != null ? v.trim() : ""; }

    public String getPersistencia() { return persistencia; }
    public void setPersistencia(String v) { this.persistencia = v != null ? v.trim() : ""; }

    public String getEnriquecimento() { return enriquecimento; }
    public void setEnriquecimento(String v) { this.enriquecimento = v != null ? v.trim() : ""; }

    public String getMapaAtributo() { return mapaAtributo; }
    public void setMapaAtributo(String v) { this.mapaAtributo = v != null ? v.trim() : ""; }

    public String getSaida() { return saida; }
    public void setSaida(String v) { this.saida = v != null ? v.trim() : ""; }

    public String getCampoConcatenado() { return campoConcatenado; }
    public void setCampoConcatenado(String v) { this.campoConcatenado = v != null ? v.trim() : ""; }

    public Integer getIdentificadorCampo() { return identificadorCampo; }
    public void setIdentificadorCampo(Integer v) { this.identificadorCampo = v; }

    public String getNomeCampo() { return nomeCampo; }
    public void setNomeCampo(String v) { this.nomeCampo = v != null ? v.trim() : ""; }

    public String getDescricaoCampo() { return descricaoCampo; }
    public void setDescricaoCampo(String v) { this.descricaoCampo = v != null ? v.trim() : ""; }

    public String getTipoCampo() { return tipoCampo; }
    public void setTipoCampo(String v) { this.tipoCampo = v != null ? v.trim() : ""; }

    public Integer getTamanhoCampo() { return tamanhoCampo; }
    public void setTamanhoCampo(Integer v) { this.tamanhoCampo = v; }

    public Integer getPosicaoInicial() { return posicaoInicial; }
    public void setPosicaoInicial(Integer v) { this.posicaoInicial = v; }

    public Integer getPosicaoFinal() { return posicaoFinal; }
    public void setPosicaoFinal(Integer v) { this.posicaoFinal = v; }

    public String getValorPadrao() { return valorPadrao; }
    public void setValorPadrao(String v) { this.valorPadrao = v != null ? v.trim() : ""; }

    public String getAlinhamentoCampo() { return alinhamentoCampo; }
    public void setAlinhamentoCampo(String v) { this.alinhamentoCampo = v != null ? v.trim() : ""; }

    public String getCampoObrigatorio() { return campoObrigatorio; }
    public void setCampoObrigatorio(String v) { this.campoObrigatorio = v != null ? v.trim() : ""; }

    public String getDominioCampo() { return dominioCampo; }
    public void setDominioCampo(String v) { this.dominioCampo = v != null ? v.trim() : ""; }

    public String getMascaraCampo() { return mascaraCampo; }
    public void setMascaraCampo(String v) { this.mascaraCampo = v != null ? v.trim() : ""; }

    public String getNomeTabela() { return nomeTabela; }
    public void setNomeTabela(String v) { this.nomeTabela = v != null ? v.trim() : ""; }

    public String getNomeColuna() { return nomeColuna; }
    public void setNomeColuna(String v) { this.nomeColuna = v != null ? v.trim() : ""; }

    public String getOracleDataType() { return oracleDataType; }
    public void setOracleDataType(String v) { this.oracleDataType = v != null ? v.trim() : ""; }

    public Integer getDataLength() { return dataLength; }
    public void setDataLength(Integer v) { this.dataLength = v; }

    public Integer getNumberPrecision() { return numberPrecision; }
    public void setNumberPrecision(Integer v) { this.numberPrecision = v; }

    public Integer getNumberScale() { return numberScale; }
    public void setNumberScale(Integer v) { this.numberScale = v; }

    public String getNullable() { return nullable; }
    public void setNullable(String v) { this.nullable = v != null ? v.trim() : ""; }

    public String getEncrypted() { return encrypted; }
    public void setEncrypted(String v) { this.encrypted = v != null ? v.trim() : ""; }

    public String getUnique() { return unique; }
    public void setUnique(String v) { this.unique = v != null ? v.trim() : ""; }

    public String getRuleAttribute() { return ruleAttribute; }
    public void setRuleAttribute(String v) { this.ruleAttribute = v != null ? v.trim() : ""; }

    public String getDefaultValue() { return defaultValue; }
    public void setDefaultValue(String v) { this.defaultValue = v != null ? v.trim() : ""; }

    public String getDescription() { return description; }
    public void setDescription(String v) { this.description = v != null ? v.trim() : ""; }

    public String getOrigin() { return origin; }
    public void setOrigin(String v) { this.origin = v != null ? v.trim() : ""; }

    public String getEventAttribute() { return eventAttribute; }
    public void setEventAttribute(String v) { this.eventAttribute = v != null ? v.trim() : ""; }

    public String getType() { return type; }
    public void setType(String v) { this.type = v != null ? v.trim() : ""; }

    public String getModelAttribute() { return modelAttribute; }
    public void setModelAttribute(String v) { this.modelAttribute = v != null ? v.trim() : ""; }

    public String getScoreModelIn() { return scoreModelIn; }
    public void setScoreModelIn(String v) { this.scoreModelIn = v != null ? v.trim() : ""; }

    public String getValorUsuario() { return valorUsuario; }
    public void setValorUsuario(String v) { this.valorUsuario = v != null ? v : ""; }

    /**
     * Calcula a posição final esperada com base na posição inicial e tamanho.
     */
    public int calcularPosicaoFinalEsperada() {
        if (posicaoInicial != null && tamanhoCampo != null) {
            return posicaoInicial + tamanhoCampo - 1;
        }
        return posicaoFinal != null ? posicaoFinal : 0;
    }

    /**
     * Retorna true se o campo é de entrada (Entrada = S) e tem posição definida.
     */
    public boolean isEntradaComPosicao() {
        return "S".equalsIgnoreCase(entrada) && posicaoInicial != null && tamanhoCampo != null;
    }

    @Override
    public String toString() {
        return String.format("[%d] %s (pos %d-%d, tam %d, tipo %s)",
                identificadorCampo, nomeCampo, posicaoInicial, posicaoFinal, tamanhoCampo, tipoCampo);
    }
}
