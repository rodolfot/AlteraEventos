"""
Gerador de XML a partir de Planilhas Excel/CSV
Interface Tkinter completa - Python 3.8+

Funcionalidades:
  - Carregar planilha principal (aba "Campos Entrada")
  - Carregar planilha origem e copiar campos para a principal
  - Tabela editável com Campo/PosIni/PosFin/Tamanho/Tipo/Alinhamento/Valor
  - Validação da soma de tamanhos x posições
  - Geração de XML com atributos posicionados corretamente
  - Salvar planilha atualizada de volta em .xlsx ou .csv
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from xml.dom import minidom
import openpyxl
import csv
import os
import re
import threading


# ─────────────────────────────────────────────────────────────────────────────
# Constantes de cor / estilo
# ─────────────────────────────────────────────────────────────────────────────
COR_BG          = "#f4f6f8"
COR_TOOLBAR     = "#eceff1"
COR_BTN_VERDE   = "#43a047"
COR_BTN_AZUL    = "#1e88e5"
COR_BTN_LARANJA = "#fb8c00"
COR_BTN_ROXO    = "#8e24aa"
COR_BTN_CINZA   = "#546e7a"
COR_BTN_TEAL    = "#00897b"
COR_BTN_VERM    = "#e53935"
COR_BRANCO      = "#ffffff"
FONT_NORMAL     = ("Segoe UI", 9)
FONT_BOLD       = ("Segoe UI", 9, "bold")
FONT_MONO       = ("Consolas", 9)


# ─────────────────────────────────────────────────────────────────────────────
# Utilitários de leitura de planilha
# ─────────────────────────────────────────────────────────────────────────────

def _normalizar_chave(texto):
    """Normaliza texto para comparação: minúsculo, sem espaços/underscores."""
    return re.sub(r"[\s_\-]", "", str(texto or "")).lower()


def _cell_str(valor, default=""):
    """Converte valor de célula para string limpa."""
    if valor is None:
        return default
    if isinstance(valor, float):
        return str(int(valor)) if valor == int(valor) else str(valor)
    return str(valor).strip()


def _cell_int(valor):
    """Converte valor de célula para int ou None."""
    try:
        return int(float(str(valor).strip()))
    except (ValueError, TypeError):
        return None


def _detectar_linha_cabecalho(sheet):
    """Localiza a linha do cabeçalho procurando 'NomeCampo' nas 10 primeiras linhas."""
    alvo = {"nomecampo", "nome", "campo", "fieldname"}
    for row_idx in range(1, min(11, sheet.max_row + 1)):
        for cell in sheet[row_idx]:
            if _normalizar_chave(cell.value) in alvo:
                return row_idx
    return 2  # padrão: linha 2


def _mapear_colunas(sheet, header_row):
    """Retorna {chave_normalizada: índice_coluna_1based} da linha de cabeçalho."""
    mapa = {}
    for cell in sheet[header_row]:
        if cell.value:
            mapa[_normalizar_chave(cell.value)] = cell.column
    return mapa


def _get_col(row_cells, col_map, *chaves, default=""):
    """Lê célula pelo nome normalizado da coluna."""
    for chave in chaves:
        col = col_map.get(_normalizar_chave(chave))
        if col:
            val = _cell_str(row_cells[col - 1].value, default)
            if val != default or default != "":
                return val
    return default


def ler_campos_entrada(filepath):
    """
    Lê a aba 'Campos Entrada' de um .xlsx ou qualquer aba de um .csv.
    Retorna lista de dicts com os campos do evento.
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".xlsx", ".xls"):
        return _ler_xlsx_campos_entrada(filepath)
    elif ext == ".csv":
        return _ler_csv_campos_entrada(filepath)
    else:
        raise ValueError(f"Formato não suportado: {ext}")


def _ler_campos_de_sheet(sheet):
    """
    Lê campos de qualquer aba de planilha, detectando cabeçalho e colunas automaticamente.
    Retorna (campos, headers) onde:
      - campos: lista de dicts com chaves padrão + '_raw' (todos os valores pelo header original)
      - headers: lista de nomes de colunas na ordem original da planilha
    """
    header_row = _detectar_linha_cabecalho(sheet)
    col_map = _mapear_colunas(sheet, header_row)

    # Cabeçalhos na ordem original da planilha
    headers = []
    col_para_header = {}
    for cell in sheet[header_row]:
        if cell.value:
            h = str(cell.value).strip()
            headers.append(h)
            col_para_header[cell.column] = h

    campos = []
    for row_idx in range(header_row + 1, sheet.max_row + 1):
        row = sheet[row_idx]

        nome = _get_col(row, col_map, "NomeCampo", "Nome", "Campo")
        if not nome:
            continue

        tamanho = _cell_int(_get_col(row, col_map, "TamanhoCampo", "Tamanho"))
        pos_ini = _cell_int(_get_col(row, col_map, "PosicaoInicial", "PosInicial", "PosIni"))
        pos_fin_lido = _cell_int(_get_col(row, col_map, "PosicaoFinal", "PosFinal", "PosFin"))
        pos_fin = (pos_ini + tamanho - 1) if (pos_ini and tamanho) else pos_fin_lido

        # Todos os valores da linha pelo nome original do cabeçalho
        raw = {}
        for cell in row:
            if cell.column in col_para_header:
                raw[col_para_header[cell.column]] = _cell_str(cell.value)

        campo = {
            "linha":      row_idx,
            "entrada":    _get_col(row, col_map, "Entrada", default="S"),
            "id":         _get_col(row, col_map, "IdentificadorCampo", "ID", "Id"),
            "nome":       nome,
            "descricao":  _get_col(row, col_map, "DescricaoCampo", "Descricao"),
            "tipo":       _get_col(row, col_map, "TipoCampo", "Tipo", default="TEXTO"),
            "tamanho":    tamanho,
            "pos_ini":    pos_ini,
            "pos_fin":    pos_fin,
            "valor_padrao": _get_col(row, col_map, "ValorPadrao", "Valor_Padrao"),
            "alinhamento":  _get_col(row, col_map, "AlinhamentoCampo", "Alinhamento"),
            "obrigatorio":  _get_col(row, col_map, "CampoObrigatorio", "Obrigatorio"),
            "coluna_db":    _get_col(row, col_map, "NomeColuna", "Coluna_DB", "ColunaDB"),
            "oracle_type":  _get_col(row, col_map, "OracleDataType", "OracleType"),
            "valor":        _get_col(row, col_map, "ValorPadrao", "Valor_Padrao"),
            "_raw":         raw,
        }
        campos.append(campo)

    return campos, headers


def _ler_xlsx_campos_entrada(filepath):
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # Localiza a aba
    sheet = None
    for name in wb.sheetnames:
        if _normalizar_chave(name) in ("camposentrada", "campos_entrada", "campos entrada"):
            sheet = wb[name]
            break
    if sheet is None:
        raise ValueError(
            f"Aba 'Campos Entrada' não encontrada.\n"
            f"Abas disponíveis: {', '.join(wb.sheetnames)}"
        )

    campos, _ = _ler_campos_de_sheet(sheet)
    return campos


def _ler_csv_campos_entrada(filepath):
    campos = []
    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for i, row in enumerate(reader):
            norm = {_normalizar_chave(k): v.strip() for k, v in row.items()}
            nome = norm.get("nomecampo") or norm.get("nome") or ""
            if not nome:
                continue

            tamanho = _cell_int(norm.get("tamanhocampo") or norm.get("tamanho"))
            pos_ini = _cell_int(norm.get("posicaoinicial") or norm.get("posinicial"))
            pos_fin = (pos_ini + tamanho - 1) if (pos_ini and tamanho) else None

            campo = {
                "linha":      i + 2,
                "entrada":    norm.get("entrada", "S"),
                "id":         norm.get("identificadorcampo") or norm.get("id", ""),
                "nome":       nome,
                "descricao":  norm.get("descricaocampo") or norm.get("descricao", ""),
                "tipo":       norm.get("tipocampo") or norm.get("tipo", "TEXTO"),
                "tamanho":    tamanho,
                "pos_ini":    pos_ini,
                "pos_fin":    pos_fin,
                "valor_padrao": norm.get("valorpadrao", ""),
                "alinhamento":  norm.get("alinhamentocampo") or norm.get("alinhamento", ""),
                "obrigatorio":  norm.get("campoobrigatorio") or norm.get("obrigatorio", ""),
                "coluna_db":    norm.get("nomecoluna") or norm.get("colunadb", ""),
                "oracle_type":  norm.get("oracledatatype", ""),
                "valor":        norm.get("valorpadrao", ""),
            }
            campos.append(campo)
    return campos


def ler_todas_abas(filepath):
    """
    Lê todas as abas de um .xlsx como dicionário {nome_aba: {"campos": list, "headers": list}}.
    Para CSV, retorna uma única entrada 'Campos Entrada'.
    Abas sem campos reconhecidos são ignoradas.
    """
    ext = os.path.splitext(filepath)[1].lower()

    if ext == ".csv":
        campos = _ler_csv_campos_entrada(filepath)
        return {"Campos Entrada": {"campos": campos, "headers": []}}

    wb = openpyxl.load_workbook(filepath, data_only=True)
    resultado = {}

    for nome_aba in wb.sheetnames:
        ws = wb[nome_aba]
        try:
            campos, headers = _ler_campos_de_sheet(ws)
            if campos:
                resultado[nome_aba] = {"campos": campos, "headers": headers}
        except Exception:
            pass

    return resultado


def _nome_xml_para_aba(nome_aba):
    """Converte nome de aba em identificador XML. Ex: 'Campos Entrada' → 'XML_ENTRADA'."""
    n = re.sub(r"^[Cc]ampos\s+", "", nome_aba).strip()
    n = re.sub(r"[\s_\-]+", "_", n).upper()
    return f"XML_{n}"


def salvar_xlsx(filepath, campos, nome_aba="Campos Entrada"):
    """Salva lista de campos na aba indicada do arquivo Excel."""
    try:
        wb = openpyxl.load_workbook(filepath)
    except Exception:
        wb = openpyxl.Workbook()
        wb.active.title = nome_aba

    sheet_name = nome_aba
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
    else:
        ws = wb[sheet_name]

    # Cabeçalho na linha 2 (padrão do template)
    headers = [
        "Entrada", "Persistencia", "Enriquecimento", "MapaAtributo", "Saida", "CampoConcatenado",
        "IdentificadorCampo", "NomeCampo", "DescricaoCampo", "TipoCampo", "TamanhoCampo",
        "PosicaoInicial", "PosicaoFinal", "ValorPadrao", "AlinhamentoCampo", "CampoObrigatorio",
    ]
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=2, column=col_idx, value=header)

    # Dados a partir da linha 6
    for i, campo in enumerate(campos):
        r = 6 + i
        ws.cell(r, 1,  campo.get("entrada", "S"))
        ws.cell(r, 7,  campo.get("id") or i + 1)
        ws.cell(r, 8,  campo.get("nome", ""))
        ws.cell(r, 9,  campo.get("descricao", ""))
        ws.cell(r, 10, campo.get("tipo", "TEXTO"))
        ws.cell(r, 11, campo.get("tamanho"))
        ws.cell(r, 12, campo.get("pos_ini"))
        if campo.get("pos_ini") and campo.get("tamanho"):
            ws.cell(r, 13).value = f"=L{r}+K{r}-1"
        ws.cell(r, 14, campo.get("valor_padrao", ""))
        ws.cell(r, 15, campo.get("alinhamento", ""))
        ws.cell(r, 16, campo.get("obrigatorio", ""))

    wb.save(filepath)


def salvar_csv(filepath, campos):
    """Salva lista de campos em CSV."""
    fieldnames = ["entrada", "id", "nome", "descricao", "tipo", "tamanho",
                  "pos_ini", "pos_fin", "valor_padrao", "alinhamento",
                  "obrigatorio", "coluna_db", "valor"]
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(campos)


# ─────────────────────────────────────────────────────────────────────────────
# XML
# ─────────────────────────────────────────────────────────────────────────────

def _sanitizar_xml(nome):
    s = re.sub(r"[^a-zA-Z0-9_\-.]", "_", (nome or "campo").strip())
    if s and not (s[0].isalpha() or s[0] == "_"):
        s = "_" + s
    return s or "campo"


def _aplicar_alinhamento(valor, tamanho, alinhamento, tipo):
    valor = str(valor or "")
    if len(valor) > tamanho:
        return valor[:tamanho]
    diff = tamanho - len(valor)
    if diff == 0:
        return valor

    alin = (alinhamento or "").upper().strip()
    if not alin:
        numerico = any(t in (tipo or "").upper() for t in ("INTEIRO", "DECIMAL", "NUMERO"))
        alin = "ZERO_ESQUERDA" if numerico else "BRANCO_ESQUERDA"

    if alin == "BRANCO_DIREITA":   return " " * diff + valor
    if alin == "BRANCO_ESQUERDA":  return valor + " " * diff
    if alin == "ZERO_ESQUERDA":    return "0" * diff + valor
    if alin == "ZERO_DIREITA":     return valor + "0" * diff
    return valor + " " * diff


def construir_xml(campos, headers=None):
    """
    Constrói string XML formatada a partir da lista de campos.
    Se 'headers' for fornecido, usa os nomes originais das colunas da aba como tags XML,
    respeitando a ordem e as colunas próprias de cada aba.
    """
    _NOME_COLS = {"nomecampo", "nome", "campo", "fieldname"}

    ativos = sorted(
        [c for c in campos if (c.get("entrada", "S") or "S").upper() == "S" and c.get("pos_ini")],
        key=lambda c: c.get("pos_ini", 0)
    )
    total = sum(c.get("tamanho", 0) or 0 for c in ativos)

    root = ET.Element("evento")
    root.set("tamanhoTotal", str(total))
    root.set("totalCampos", str(len(ativos)))

    campos_el = ET.SubElement(root, "campos")

    for c in ativos:
        campo_el = ET.SubElement(campos_el, "campo")

        def _sub(tag, val, _el=campo_el):
            tag_safe = _sanitizar_xml(tag)
            if val is not None and str(val).strip():
                e = ET.SubElement(_el, tag_safe)
                e.text = str(val)

        # Primeiro elemento: NomeCampo
        _sub("NomeCampo", c.get("nome", ""))

        # Segundo elemento: valor posicionado (calculado)
        valor = c.get("valor") or c.get("valor_padrao") or ""
        if c.get("tamanho"):
            valor = _aplicar_alinhamento(valor, c["tamanho"], c.get("alinhamento", ""), c.get("tipo", ""))
        _sub("valor", valor)

        if headers and "_raw" in c:
            # Usa as colunas reais da aba, na ordem do cabeçalho original
            raw = c["_raw"]
            for header in headers:
                if _normalizar_chave(header) in _NOME_COLS:
                    continue  # NomeCampo já incluído acima
                val = raw.get(header, "")
                if val:
                    _sub(header, val)
        else:
            # Fallback: colunas fixas (campos sem _raw ou sem headers)
            if c.get("id"):           _sub("id",             str(c["id"]))
            _sub("tipo",               c.get("tipo") or "TEXTO")
            if c.get("tamanho"):      _sub("tamanho",        str(c["tamanho"]))
            if c.get("pos_ini"):      _sub("posicaoInicial", str(c["pos_ini"]))

            pos_fin = c.get("pos_fin")
            if not pos_fin and c.get("pos_ini") and c.get("tamanho"):
                pos_fin = c["pos_ini"] + c["tamanho"] - 1
            if pos_fin:               _sub("posicaoFinal",   str(pos_fin))

            if c.get("alinhamento"):  _sub("alinhamento",    c["alinhamento"])
            if c.get("obrigatorio"):  _sub("obrigatorio",    c["obrigatorio"])
            if c.get("descricao"):    _sub("descricao",      c["descricao"])
            if c.get("coluna_db"):    _sub("colunaDB",       c["coluna_db"])

    raw_xml = ET.tostring(root, encoding="unicode")
    return minidom.parseString(raw_xml).toprettyxml(indent="    ")


# ─────────────────────────────────────────────────────────────────────────────
# Validação
# ─────────────────────────────────────────────────────────────────────────────

def validar_campos(campos):
    """Retorna (erros, avisos, infos) com os resultados da validação."""
    erros, avisos, infos = [], [], []

    ativos = [
        c for c in campos
        if (c.get("entrada", "S") or "S").upper() == "S"
        and c.get("pos_ini") and c.get("tamanho")
    ]

    if not ativos:
        return erros, ["Nenhum campo ativo com posição definida."], infos

    ordenados = sorted(ativos, key=lambda c: c["pos_ini"])
    total = sum(c["tamanho"] for c in ordenados)

    # 1. Fórmula PosicaoFinal
    for c in ordenados:
        esperado = c["pos_ini"] + c["tamanho"] - 1
        if c.get("pos_fin") and c["pos_fin"] != esperado:
            erros.append(
                f"Campo '{c['nome']}': PosicaoFinal={c['pos_fin']} "
                f"mas esperado {esperado} (PosIni={c['pos_ini']} + Tam={c['tamanho']} - 1)"
            )

    # 2. Começa em 1
    if ordenados[0]["pos_ini"] != 1:
        avisos.append(
            f"Layout não começa em 1. Primeiro campo '{ordenados[0]['nome']}' "
            f"inicia em {ordenados[0]['pos_ini']}."
        )

    # 3. Continuidade
    for i in range(1, len(ordenados)):
        ant, atu = ordenados[i - 1], ordenados[i]
        prox_esperado = ant["pos_ini"] + ant["tamanho"]
        if atu["pos_ini"] > prox_esperado:
            avisos.append(
                f"GAP entre '{ant['nome']}' (term. {prox_esperado-1}) "
                f"e '{atu['nome']}' (inicia {atu['pos_ini']}) "
                f"— {atu['pos_ini'] - prox_esperado} byte(s)."
            )
        elif atu["pos_ini"] < prox_esperado:
            erros.append(
                f"SOBREPOSIÇÃO: '{ant['nome']}' e '{atu['nome']}' "
                f"se sobrepõem em pos={atu['pos_ini']}."
            )

    # 4. Obrigatórios sem valor
    for c in ordenados:
        if (c.get("obrigatorio") or "").upper() == "S":
            v = (c.get("valor") or c.get("valor_padrao") or "").strip()
            if not v:
                avisos.append(f"Campo obrigatório '{c['nome']}' sem valor preenchido.")

    infos.append(f"Campos de entrada: {len(ordenados)}")
    infos.append(f"Soma dos tamanhos: {total} bytes")
    if ordenados:
        infos.append(f"Posição final do layout: {ordenados[-1]['pos_ini'] + ordenados[-1]['tamanho'] - 1}")

    return erros, avisos, infos


# ─────────────────────────────────────────────────────────────────────────────
# Janela de carregamento (loading)
# ─────────────────────────────────────────────────────────────────────────────

class JanelaCarregando(tk.Toplevel):
    """Diálogo modal com barra de progresso indeterminada exibido durante carregamento."""

    def __init__(self, parent, mensagem="Carregando planilha..."):
        super().__init__(parent)
        self.title("Aguarde")
        self.resizable(False, False)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", lambda: None)  # impede fechar

        frm = tk.Frame(self, bg=COR_BG, padx=40, pady=28)
        frm.pack(fill=tk.BOTH, expand=True)

        tk.Label(frm, text="⏳", bg=COR_BG, font=("Segoe UI", 32)).pack()
        tk.Label(
            frm, text=mensagem, bg=COR_BG,
            font=FONT_BOLD, fg="#333333", wraplength=300
        ).pack(pady=(10, 14))

        self._bar = ttk.Progressbar(frm, mode="indeterminate", length=280)
        self._bar.pack()
        self._bar.start(12)

        # Centraliza sobre o parent
        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width() // 2 - self.winfo_width() // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{px}+{py}")

        self.grab_set()

    def fechar(self):
        try:
            self._bar.stop()
            self.grab_release()
            self.destroy()
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# Janela de edição de campo (Toplevel)
# ─────────────────────────────────────────────────────────────────────────────

class JanelaEditarCampo(tk.Toplevel):
    TIPOS = ["TEXTO", "INTEIRO", "INTEIRO_LONGO", "DECIMAL", "DATA", "DATA_HORA", "HORA", "BOOLEANO"]
    ALINHAMENTOS = ["", "BRANCO_ESQUERDA", "BRANCO_DIREITA", "ZERO_ESQUERDA", "ZERO_DIREITA"]
    ENTRADAS = ["S", "N"]
    OBRIGATORIOS = ["", "S", "N"]

    def __init__(self, parent, campo=None, on_confirmar=None):
        super().__init__(parent)
        self.title("Editar Campo" if campo else "Novo Campo")
        self.resizable(False, False)
        self.grab_set()

        self.on_confirmar = on_confirmar
        self.resultado = None

        self._vars = {}
        self._build_ui(campo or {})

        self.transient(parent)
        self.wait_window(self)

    def _build_ui(self, campo):
        pad = {"padx": 8, "pady": 3}

        frame = tk.Frame(self, bg=COR_BG, padx=16, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        linhas = [
            ("Entrada:",     "entrada",    ttk.Combobox, {"values": self.ENTRADAS, "width": 6}),
            ("Nome:",        "nome",       tk.Entry,     {"width": 32}),
            ("Descrição:",   "descricao",  tk.Entry,     {"width": 32}),
            ("Tipo:",        "tipo",       ttk.Combobox, {"values": self.TIPOS, "width": 18}),
            ("Tamanho:",     "tamanho",    tk.Entry,     {"width": 10}),
            ("Pos. Inicial:","pos_ini",    tk.Entry,     {"width": 10}),
            ("Pos. Final:",  "pos_fin",    tk.Entry,     {"width": 10}),
            ("Alinhamento:", "alinhamento",ttk.Combobox, {"values": self.ALINHAMENTOS, "width": 20}),
            ("Obrigatório:", "obrigatorio",ttk.Combobox, {"values": self.OBRIGATORIOS, "width": 6}),
            ("Valor Padrão:","valor_padrao",tk.Entry,    {"width": 24}),
            ("Valor (XML):", "valor",      tk.Entry,     {"width": 24}),
            ("Coluna DB:",   "coluna_db",  tk.Entry,     {"width": 24}),
        ]

        for i, (label, key, WidgetClass, kw) in enumerate(linhas):
            tk.Label(frame, text=label, bg=COR_BG, font=FONT_NORMAL,
                     anchor=tk.W).grid(row=i, column=0, sticky=tk.W, **pad)

            var = tk.StringVar()
            self._vars[key] = var

            w_kw = dict(kw)
            if WidgetClass == ttk.Combobox:
                w = WidgetClass(frame, textvariable=var, **w_kw)
                w.configure(state="readonly" if key in ("entrada", "tipo", "alinhamento", "obrigatorio") else "normal")
            else:
                w = WidgetClass(frame, textvariable=var, font=FONT_NORMAL, **w_kw)

            # Pos Final read-only (calculado)
            if key == "pos_fin":
                w.configure(state="readonly")

            w.grid(row=i, column=1, sticky=tk.W, **pad)

        frame.columnconfigure(1, weight=1)

        # Auto-calcular Pos Final
        for k in ("tamanho", "pos_ini"):
            self._vars[k].trace_add("write", self._calc_pos_fin)

        # Preenche com valores existentes
        mapa_default = {"entrada": "S", "tipo": "TEXTO"}
        for key, var in self._vars.items():
            val = campo.get(key)
            if val is not None and str(val).strip():
                var.set(str(val))
            elif key in mapa_default:
                var.set(mapa_default[key])

        # Botões
        btn_frame = tk.Frame(self, bg=COR_BG, pady=8)
        btn_frame.pack()

        btn_ok = tk.Button(btn_frame, text="Confirmar", font=FONT_BOLD,
                           bg=COR_BTN_VERDE, fg=COR_BRANCO, relief=tk.FLAT,
                           padx=12, pady=4, command=self._confirmar)
        btn_ok.pack(side=tk.LEFT, padx=6)

        btn_cancel = tk.Button(btn_frame, text="Cancelar", font=FONT_NORMAL,
                               bg=COR_BTN_CINZA, fg=COR_BRANCO, relief=tk.FLAT,
                               padx=12, pady=4, command=self.destroy)
        btn_cancel.pack(side=tk.LEFT, padx=6)

        self.bind("<Return>", lambda e: self._confirmar())
        self.bind("<Escape>", lambda e: self.destroy())

    def _calc_pos_fin(self, *_):
        try:
            pos = int(self._vars["pos_ini"].get())
            tam = int(self._vars["tamanho"].get())
            self._vars["pos_fin"].set(str(pos + tam - 1))
        except ValueError:
            pass

    def _confirmar(self):
        nome = self._vars["nome"].get().strip()
        if not nome:
            messagebox.showwarning("Aviso", "Nome do campo é obrigatório.", parent=self)
            return

        try:
            tamanho = int(self._vars["tamanho"].get()) if self._vars["tamanho"].get() else None
            pos_ini = int(self._vars["pos_ini"].get()) if self._vars["pos_ini"].get() else None
        except ValueError:
            messagebox.showerror("Erro", "Tamanho e Pos. Inicial devem ser números inteiros.", parent=self)
            return

        self.resultado = {
            "entrada":    self._vars["entrada"].get() or "S",
            "nome":       nome,
            "descricao":  self._vars["descricao"].get().strip(),
            "tipo":       self._vars["tipo"].get() or "TEXTO",
            "tamanho":    tamanho,
            "pos_ini":    pos_ini,
            "pos_fin":    (pos_ini + tamanho - 1) if (pos_ini and tamanho) else None,
            "alinhamento":  self._vars["alinhamento"].get(),
            "obrigatorio":  self._vars["obrigatorio"].get(),
            "valor_padrao": self._vars["valor_padrao"].get().strip(),
            "valor":        self._vars["valor"].get(),
            "coluna_db":    self._vars["coluna_db"].get().strip(),
        }

        if self.on_confirmar:
            self.on_confirmar(self.resultado)
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Janela de seleção múltipla de campos da origem
# ─────────────────────────────────────────────────────────────────────────────

class JanelaCopiarCampos(tk.Toplevel):
    """
    Diálogo de seleção de campos da planilha origem.
    Exibe uma aba por sheet (estilo Excel) com Listbox de campos em cada uma.
    Aceita dados_por_aba: {nome_aba: {"campos": list, "headers": list}}
    """

    def __init__(self, parent, dados_por_aba, on_confirmar=None):
        super().__init__(parent)
        self.title("Copiar Campos da Origem")
        self.geometry("520x560")
        self.resizable(True, True)
        self.minsize(400, 400)
        self.grab_set()
        self.transient(parent)

        self._dados_por_aba = dados_por_aba
        self._aba_ativa = ""
        self._listboxes: dict = {}          # nome_aba → Listbox
        self._campos_filtrados: dict = {}   # nome_aba → lista filtrada
        self.on_confirmar = on_confirmar

        self._build_ui()
        self.wait_window(self)

    def _build_ui(self):
        frame = tk.Frame(self, bg=COR_BG, padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Filtro
        filt_frame = tk.Frame(frame, bg=COR_BG)
        filt_frame.pack(fill=tk.X, pady=(0, 6))
        tk.Label(filt_frame, text="🔍 Filtrar:", bg=COR_BG, font=FONT_NORMAL).pack(side=tk.LEFT)
        self._var_filtro = tk.StringVar()
        self._var_filtro.trace_add("write", lambda *_: self._filtrar())
        tk.Entry(filt_frame, textvariable=self._var_filtro, font=FONT_NORMAL, width=28,
                 relief=tk.FLAT, highlightthickness=1,
                 highlightbackground="#bbb").pack(side=tk.LEFT, padx=4)
        tk.Button(filt_frame, text="✕", command=lambda: self._var_filtro.set(""),
                  bg="#ddd", relief=tk.FLAT, font=FONT_NORMAL).pack(side=tk.LEFT)

        tk.Label(frame, text="Selecione os campos (Ctrl+Click ou Shift+Click para múltiplos):",
                 bg=COR_BG, font=FONT_NORMAL, anchor=tk.W).pack(fill=tk.X, pady=(0, 4))

        # Notebook com uma aba por sheet
        self._nb = ttk.Notebook(frame)
        self._nb.pack(fill=tk.BOTH, expand=True)
        self._nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        for nome_aba, aba_info in self._dados_por_aba.items():
            tab_frame = tk.Frame(self._nb, bg=COR_BG, padx=4, pady=4)
            self._nb.add(tab_frame, text=f"  {nome_aba}  ")
            lb = self._criar_listbox(tab_frame)
            self._listboxes[nome_aba] = lb
            self._campos_filtrados[nome_aba] = list(aba_info.get("campos", []))

        # Seleciona a primeira aba
        nomes = list(self._dados_por_aba.keys())
        if nomes:
            self._aba_ativa = nomes[0]
            self._nb.select(0)

        # Botões de seleção rápida + contador
        sel_frame = tk.Frame(frame, bg=COR_BG)
        sel_frame.pack(fill=tk.X, pady=(6, 0))
        tk.Button(sel_frame, text="Selecionar Todos", command=self._sel_todos,
                  bg="#eceff1", font=FONT_NORMAL, relief=tk.FLAT).pack(side=tk.LEFT, padx=2)
        tk.Button(sel_frame, text="Limpar Seleção", command=self._sel_nenhum,
                  bg="#eceff1", font=FONT_NORMAL, relief=tk.FLAT).pack(side=tk.LEFT, padx=2)
        self._lbl_sel = tk.Label(sel_frame, text="0 selecionado(s)",
                                  bg=COR_BG, fg="#555", font=FONT_NORMAL)
        self._lbl_sel.pack(side=tk.RIGHT)

        # Botões de ação
        btn_frame = tk.Frame(self, bg=COR_BG, pady=8)
        btn_frame.pack()
        self._btn_copiar = tk.Button(btn_frame, text="⬇ Copiar 0 campos", font=FONT_BOLD,
                                      bg=COR_BTN_LARANJA, fg=COR_BRANCO, relief=tk.FLAT,
                                      padx=12, pady=4, command=self._confirmar,
                                      state=tk.DISABLED)
        self._btn_copiar.pack(side=tk.LEFT, padx=6)
        tk.Button(btn_frame, text="Cancelar", font=FONT_NORMAL,
                  bg=COR_BTN_CINZA, fg=COR_BRANCO, relief=tk.FLAT,
                  padx=12, pady=4, command=self.destroy).pack(side=tk.LEFT, padx=6)

        self.bind("<Escape>", lambda e: self.destroy())
        self._filtrar()

    def _criar_listbox(self, parent):
        """Cria e retorna um Listbox com scrollbar dentro do parent."""
        list_frame = tk.Frame(parent, bg=COR_BG)
        list_frame.pack(fill=tk.BOTH, expand=True)
        lb = tk.Listbox(list_frame, selectmode=tk.EXTENDED, font=FONT_NORMAL,
                        activestyle="dotbox", exportselection=False,
                        relief=tk.FLAT, highlightthickness=1, highlightbackground="#bbb")
        vsb = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=lb.yview)
        lb.configure(yscrollcommand=vsb.set)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        lb.bind("<<ListboxSelect>>", self._on_select)
        return lb

    def _on_tab_changed(self, _evt=None):
        """Atualiza aba ativa e reseta o contador de seleção."""
        nomes = list(self._dados_por_aba.keys())
        tabs = self._nb.tabs()
        if not tabs:
            return
        idx = self._nb.index("current")
        if idx < len(nomes):
            self._aba_ativa = nomes[idx]
        self._on_select()

    def _get_lb(self):
        """Retorna o Listbox da aba atualmente selecionada."""
        return self._listboxes.get(self._aba_ativa)

    def _filtrar(self):
        """Aplica filtro de texto em todas as abas."""
        filtro = self._var_filtro.get().lower()
        for nome_aba, aba_info in self._dados_por_aba.items():
            campos = aba_info.get("campos", [])
            filtrados = [
                c for c in campos
                if not filtro
                   or filtro in (c.get("nome") or "").lower()
                   or filtro in (c.get("descricao") or "").lower()
            ]
            self._campos_filtrados[nome_aba] = filtrados
            lb = self._listboxes.get(nome_aba)
            if not lb:
                continue
            lb.delete(0, tk.END)
            for c in filtrados:
                nome   = c.get("nome", "")
                extras = []
                if c.get("tipo"):    extras.append(c["tipo"])
                if c.get("tamanho"): extras.append(f"{c['tamanho']}b")
                label = f"{nome}  [{', '.join(extras)}]" if extras else nome
                lb.insert(tk.END, label)
        self._on_select()

    def _sel_todos(self):
        lb = self._get_lb()
        if lb:
            lb.select_set(0, tk.END)
        self._on_select()

    def _sel_nenhum(self):
        lb = self._get_lb()
        if lb:
            lb.selection_clear(0, tk.END)
        self._on_select()

    def _on_select(self, _evt=None):
        lb = self._get_lb()
        n = len(lb.curselection()) if lb else 0
        self._lbl_sel.config(text=f"{n} selecionado(s)")
        plural = "s" if n != 1 else ""
        self._btn_copiar.config(
            text=f"⬇ Copiar {n} campo{plural}",
            state=tk.NORMAL if n > 0 else tk.DISABLED
        )

    def _confirmar(self):
        lb = self._get_lb()
        if not lb:
            self.destroy()
            return
        indices = lb.curselection()
        filtrados = self._campos_filtrados.get(self._aba_ativa, [])
        resultado = [filtrados[i] for i in indices]
        if self.on_confirmar:
            self.on_confirmar(resultado)
        self.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# Aplicação principal
# ─────────────────────────────────────────────────────────────────────────────

class GeradorXMLApp:

    # ── Inicialização ────────────────────────────────────────────────────────

    def __init__(self, root):
        self.root = root
        self.root.title("Gerador de XML — Planilhas de Eventos")
        self.root.geometry("1280x820")
        self.root.configure(bg=COR_BG)
        self.root.minsize(900, 600)

        self._campos: list = []             # campos da aba ativa
        self._dados_por_aba: dict = {}      # nome_aba → {"campos": list, "headers": list}
        self._aba_ativa: str = ""           # aba atualmente exibida
        self._headers_ativos: list = []     # headers da aba ativa (ordem original da planilha)
        self._dados_por_aba_origem: dict = {}  # nome_aba → {"campos": list, "headers": list}
        self._arquivo_principal = None
        self._arquivo_origem = None
        self._idx_editando = -1             # índice do campo sendo editado

        # Widgets do notebook de abas (criados em _build_tabela)
        self._nb_abas = None                # ttk.Notebook da planilha principal
        self._trees_abas: dict = {}         # nome_aba → Treeview
        self._tree = None                   # Treeview da aba atualmente ativa
        self._ignorar_tab_change = False    # evita recursão ao selecionar aba

        self._setup_estilos()
        self._build_ui()
        self._bind_atalhos()

    def _setup_estilos(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("Treeview", rowheight=24, font=FONT_NORMAL, background=COR_BRANCO)
        s.configure("Treeview.Heading", font=FONT_BOLD)
        s.map("Treeview", background=[("selected", "#bbdefb")])

    # ── Construção da UI ─────────────────────────────────────────────────────

    def _build_ui(self):
        self._build_menubar()
        self._build_toolbar()
        self._build_conteudo()
        self._build_statusbar()

    def _build_menubar(self):
        mb = tk.Menu(self.root)
        self.root.config(menu=mb)

        m_arq = tk.Menu(mb, tearoff=0)
        m_arq.add_command(label="Carregar Principal...  Ctrl+O", command=self.carregar_principal)
        m_arq.add_command(label="Carregar Origem...",             command=self.carregar_origem)
        m_arq.add_separator()
        m_arq.add_command(label="Salvar Planilha  Ctrl+S",        command=self.salvar_planilha)
        m_arq.add_separator()
        m_arq.add_command(label="Sair",                           command=self.root.quit)
        mb.add_cascade(label="Arquivo", menu=m_arq)

        m_fer = tk.Menu(mb, tearoff=0)
        m_fer.add_command(label="Validar Soma  F5",               command=self.validar)
        m_fer.add_command(label="Recalcular Posições",            command=self.recalcular_posicoes)
        m_fer.add_separator()
        m_fer.add_command(label="Gerar XML  F6",                  command=self.gerar_xml)
        m_fer.add_command(label="Pré-visualizar XML  F7",         command=self.preview_xml)
        mb.add_cascade(label="Ferramentas", menu=m_fer)

    def _btn(self, parent, text, command, bg, **pack_kw):
        b = tk.Button(parent, text=text, command=command, bg=bg,
                      fg=COR_BRANCO, relief=tk.FLAT, font=FONT_NORMAL,
                      padx=9, pady=4, cursor="hand2",
                      activebackground=bg, activeforeground=COR_BRANCO)
        b.pack(**pack_kw)
        return b

    def _build_toolbar(self):
        bar = tk.Frame(self.root, bg=COR_TOOLBAR, pady=5)
        bar.pack(fill=tk.X)

        # ── Planilha principal ───────────────────────────────────────────────
        frp = tk.LabelFrame(bar, text="Planilha Principal", bg=COR_TOOLBAR,
                            font=FONT_NORMAL, padx=4, pady=2)
        frp.pack(side=tk.LEFT, padx=8)

        self._btn(frp, "📂 Carregar Principal", self.carregar_principal,
                  COR_BTN_VERDE, side=tk.LEFT, padx=2)
        self._lbl_principal = tk.Label(frp, text="—", bg=COR_TOOLBAR,
                                       fg="#555", font=("Segoe UI", 8))
        self._lbl_principal.pack(side=tk.LEFT, padx=6)

        # ── Planilha origem ──────────────────────────────────────────────────
        fro = tk.LabelFrame(bar, text="Planilha Origem", bg=COR_TOOLBAR,
                            font=FONT_NORMAL, padx=4, pady=2)
        fro.pack(side=tk.LEFT, padx=4)

        self._btn(fro, "📂 Carregar Origem", self.carregar_origem,
                  COR_BTN_AZUL, side=tk.LEFT, padx=2)
        self._lbl_origem = tk.Label(fro, text="—", bg=COR_TOOLBAR,
                                    fg="#555", font=("Segoe UI", 8))
        self._lbl_origem.pack(side=tk.LEFT, padx=6)

        # ── Copiar campos ─────────────────────────────────────────────────────
        frc = tk.LabelFrame(bar, text="Copiar Campos da Origem", bg=COR_TOOLBAR,
                            font=FONT_NORMAL, padx=4, pady=2)
        frc.pack(side=tk.LEFT, padx=4)

        self._btn_copiar_origem = self._btn(frc, "⬇ Copiar Campos...", self.copiar_campo,
                  COR_BTN_LARANJA, side=tk.LEFT, padx=2)
        self._btn_copiar_origem.configure(state=tk.DISABLED)

        # ── Ações ────────────────────────────────────────────────────────────
        fra = tk.Frame(bar, bg=COR_TOOLBAR)
        fra.pack(side=tk.RIGHT, padx=8)

        self._btn(fra, "📄 Gerar XML [F6]",        self.gerar_xml,         COR_BTN_ROXO,  side=tk.RIGHT, padx=2)
        self._btn(fra, "👁 Preview XML [F7]",      self.preview_xml,       COR_BTN_CINZA, side=tk.RIGHT, padx=2)
        self._btn(fra, "✔ Validar [F5]",           self.validar,           COR_BTN_CINZA, side=tk.RIGHT, padx=2)
        self._btn(fra, "💾 Salvar Planilha",        self.salvar_planilha,   COR_BTN_TEAL,  side=tk.RIGHT, padx=2)

    def _build_conteudo(self):
        paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL,
                               sashwidth=6, bg="#cfd8dc")
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=4)

        # ── Esquerda: tabela ─────────────────────────────────────────────────
        frame_esq = tk.Frame(paned, bg=COR_BG)
        paned.add(frame_esq, width=810)
        self._build_tabela(frame_esq)

        # ── Direita: detalhes / validação / XML ──────────────────────────────
        frame_dir = tk.Frame(paned, bg=COR_BG)
        paned.add(frame_dir, width=400)
        self._build_painel_direito(frame_dir)

    def _build_tabela(self, parent):
        # Barra de filtro
        filt = tk.Frame(parent, bg=COR_BG)
        filt.pack(fill=tk.X, pady=(2, 4))

        tk.Label(filt, text="🔍 Filtrar:", bg=COR_BG, font=FONT_NORMAL).pack(side=tk.LEFT)
        self._var_filtro = tk.StringVar()
        self._var_filtro.trace_add("write", lambda *_: self._atualizar_tabela())
        tk.Entry(filt, textvariable=self._var_filtro, font=FONT_NORMAL, width=28,
                 relief=tk.FLAT, highlightthickness=1,
                 highlightbackground="#bbb").pack(side=tk.LEFT, padx=4)
        tk.Button(filt, text="✕", command=lambda: self._var_filtro.set(""),
                  bg="#ddd", relief=tk.FLAT, font=FONT_NORMAL).pack(side=tk.LEFT)

        # Notebook de abas (uma por planilha)
        self._nb_abas = ttk.Notebook(parent)
        self._nb_abas.pack(fill=tk.BOTH, expand=True)
        self._nb_abas.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # Barra de ações da tabela
        bbar = tk.Frame(parent, bg=COR_BG, pady=4)
        bbar.pack(fill=tk.X)

        self._btn(bbar, "+ Novo",       self.novo_campo,       COR_BTN_VERDE,  side=tk.LEFT, padx=2)
        self._btn(bbar, "✎ Editar",    self.editar_campo,     COR_BTN_AZUL,   side=tk.LEFT, padx=2)
        self._btn(bbar, "🗑 Remover",   self.remover_campo,    COR_BTN_VERM,   side=tk.LEFT, padx=2)
        self._btn(bbar, "⟳ Recalcular",self.recalcular_posicoes, COR_BTN_LARANJA, side=tk.LEFT, padx=2)

        self._lbl_count = tk.Label(bbar, text="0 campos", bg=COR_BG,
                                   fg="#555", font=FONT_NORMAL)
        self._lbl_count.pack(side=tk.RIGHT, padx=8)

    def _build_painel_direito(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill=tk.BOTH, expand=True)

        # Aba Validação
        frv = tk.Frame(nb, bg=COR_BG, padx=4, pady=4)
        nb.add(frv, text="  Validação  ")
        self._build_aba_validacao(frv)

        # Aba XML Preview
        frx = tk.Frame(nb, bg=COR_BG, padx=4, pady=4)
        nb.add(frx, text="  XML Preview  ")
        self._build_aba_xml(frx)

        self._notebook = nb

    def _build_aba_validacao(self, parent):
        self._btn(parent, "▶ Validar Agora [F5]", self.validar,
                  COR_BTN_CINZA, anchor=tk.NW, pady=(0, 6))

        self._txt_val = tk.Text(parent, font=FONT_MONO, wrap=tk.WORD,
                                state=tk.DISABLED, bg="#fafafa",
                                relief=tk.FLAT, highlightthickness=1,
                                highlightbackground="#ccc")
        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self._txt_val.yview)
        self._txt_val.configure(yscrollcommand=vsb.set)

        self._txt_val.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Tags de cor para o texto
        self._txt_val.tag_configure("ok",     foreground="#2e7d32", font=FONT_BOLD)
        self._txt_val.tag_configure("erro",   foreground="#c62828", font=FONT_BOLD)
        self._txt_val.tag_configure("aviso",  foreground="#e65100")
        self._txt_val.tag_configure("info",   foreground="#1565c0")
        self._txt_val.tag_configure("titulo", font=FONT_BOLD)

    def _build_aba_xml(self, parent):
        self._btn(parent, "🔄 Atualizar Preview [F7]", self.preview_xml,
                  COR_BTN_ROXO, anchor=tk.NW, pady=(0, 6))

        # Sub-frame uses grid internally; parent keeps only pack
        frm = tk.Frame(parent)
        frm.pack(fill=tk.BOTH, expand=True)

        self._txt_xml = tk.Text(frm, font=FONT_MONO, wrap=tk.NONE,
                                bg="#1e1e1e", fg="#d4d4d4",
                                insertbackground="white",
                                relief=tk.FLAT)
        vsb = ttk.Scrollbar(frm, orient=tk.VERTICAL,   command=self._txt_xml.yview)
        hsb = ttk.Scrollbar(frm, orient=tk.HORIZONTAL, command=self._txt_xml.xview)
        self._txt_xml.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self._txt_xml.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

        # Syntax highlight básico
        self._txt_xml.tag_configure("tag",   foreground="#569cd6")
        self._txt_xml.tag_configure("attr",  foreground="#9cdcfe")
        self._txt_xml.tag_configure("value", foreground="#ce9178")

    def _build_statusbar(self):
        bar = tk.Frame(self.root, bg="#37474f", height=26)
        bar.pack(fill=tk.X, side=tk.BOTTOM)
        bar.pack_propagate(False)

        self._var_status = tk.StringVar(value="Pronto. Carregue uma planilha para começar.")
        tk.Label(bar, textvariable=self._var_status, bg="#37474f", fg=COR_BRANCO,
                 font=FONT_NORMAL, anchor=tk.W, padx=10).pack(side=tk.LEFT, fill=tk.X, expand=True)

        self._var_total = tk.StringVar()
        tk.Label(bar, textvariable=self._var_total, bg="#37474f", fg="#90a4ae",
                 font=FONT_NORMAL, padx=10).pack(side=tk.RIGHT)

    # ── Atalhos ──────────────────────────────────────────────────────────────

    def _bind_atalhos(self):
        self.root.bind("<Control-o>", lambda _: self.carregar_principal())
        self.root.bind("<Control-s>", lambda _: self.salvar_planilha())
        self.root.bind("<F5>",        lambda _: self.validar())
        self.root.bind("<F6>",        lambda _: self.gerar_xml())
        self.root.bind("<F7>",        lambda _: self.preview_xml())
        self.root.bind("<Delete>",    lambda _: self.remover_campo())

    # ── Carregar planilhas ────────────────────────────────────────────────────

    def _executar_em_thread(self, tarefa, on_sucesso, on_erro, mensagem="Carregando..."):
        """
        Executa `tarefa()` em thread separada enquanto exibe JanelaCarregando.
        Chama on_sucesso(resultado) ou on_erro(excecao) na thread principal via after().
        """
        janela = JanelaCarregando(self.root, mensagem)

        def _runner():
            try:
                resultado = tarefa()
                self.root.after(0, lambda: _finalizar(resultado, None))
            except Exception as exc:
                self.root.after(0, lambda e=exc: _finalizar(None, e))

        def _finalizar(resultado, erro):
            janela.fechar()
            if erro:
                on_erro(erro)
            else:
                on_sucesso(resultado)

        threading.Thread(target=_runner, daemon=True).start()

    def carregar_principal(self):
        path = filedialog.askopenfilename(
            title="Carregar Planilha Principal",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if not path:
            return

        def _tarefa():
            return ler_todas_abas(path)

        def _sucesso(dados):
            if not dados:
                messagebox.showwarning("Aviso", "Nenhuma aba com campos detectados encontrada.")
                return

            self._dados_por_aba = dados
            self._arquivo_principal = path

            nomes = list(dados.keys())
            aba_padrao = next(
                (n for n in nomes if _normalizar_chave(n) in ("camposentrada",)),
                nomes[0]
            )

            self._reconstruir_abas_principal(dados, aba_padrao)

            nome = os.path.basename(path)
            self._lbl_principal.config(text=nome, fg="#1565c0")
            total = sum(len(v["campos"]) for v in dados.values())
            self._set_status(
                f"Principal carregada: {nome}  —  {len(dados)} aba(s), {total} campos"
            )

        def _erro(e):
            messagebox.showerror("Erro ao carregar", str(e))

        self._executar_em_thread(_tarefa, _sucesso, _erro, "Carregando planilha principal...")

    def carregar_origem(self):
        path = filedialog.askopenfilename(
            title="Carregar Planilha Origem",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Todos", "*.*")]
        )
        if not path:
            return

        def _tarefa():
            dados = ler_todas_abas(path)
            if not dados:
                # Fallback: lê a primeira aba genericamente
                campos = self._ler_xlsx_generico(path)
                nome_fallback = os.path.splitext(os.path.basename(path))[0]
                dados = {nome_fallback: {"campos": campos, "headers": []}}
            return dados

        def _sucesso(dados):
            self._dados_por_aba_origem = dados
            self._arquivo_origem = path
            self._btn_copiar_origem.configure(state=tk.NORMAL)

            nome = os.path.basename(path)
            self._lbl_origem.config(text=nome, fg="#2e7d32")
            n_abas = len(dados)
            total = sum(len(v["campos"]) for v in dados.values())
            self._set_status(
                f"Origem carregada: {nome}  —  {n_abas} aba(s), {total} campos disponíveis"
            )

        def _erro(e):
            messagebox.showerror("Erro ao carregar origem", str(e))

        self._executar_em_thread(_tarefa, _sucesso, _erro, "Carregando planilha origem...")

    # ── Abas estilo Excel (notebook) ──────────────────────────────────────────

    def _criar_tree_aba(self, parent, headers):
        """Cria um Treeview com as colunas originais da aba e retorna o widget."""
        cols = headers if headers else ["NomeCampo"]

        wrap = tk.Frame(parent, bg=COR_BG)
        wrap.pack(fill=tk.BOTH, expand=True)

        tree = ttk.Treeview(wrap, columns=cols, show="headings", selectmode="browse")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor=tk.W, minwidth=50)

        vsb = ttk.Scrollbar(wrap, orient=tk.VERTICAL,   command=tree.yview)
        hsb = ttk.Scrollbar(wrap, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        tree.tag_configure("par",   background="#f5f9ff")
        tree.tag_configure("impar", background=COR_BRANCO)
        tree.tag_configure("erro",  background="#ffebee")
        tree.tag_configure("aviso", background="#fff8e1")

        tree.bind("<<TreeviewSelect>>", self._on_selecionar)
        tree.bind("<Double-1>",         self._on_duplo_clique)

        return tree

    def _reconstruir_abas_principal(self, dados, aba_padrao):
        """Recria todas as abas do Notebook com os dados carregados."""
        self._ignorar_tab_change = True

        # Limpa abas e trees anteriores
        for tab in self._nb_abas.tabs():
            self._nb_abas.forget(tab)
        self._trees_abas.clear()

        nomes = list(dados.keys())
        for nome_aba in nomes:
            aba_info = dados[nome_aba]
            headers = aba_info.get("headers", [])
            campos  = aba_info.get("campos", [])

            frame = tk.Frame(self._nb_abas, bg=COR_BG)
            self._nb_abas.add(frame, text=f"  {nome_aba}  ")

            tree = self._criar_tree_aba(frame, headers)
            self._trees_abas[nome_aba] = tree

            # Popula com os dados brutos
            for i, c in enumerate(campos):
                raw    = c.get("_raw", {})
                vals   = tuple(raw.get(h, "") for h in headers) if headers else (c.get("nome", ""),)
                pos_ini = c.get("pos_ini")
                tam     = c.get("tamanho")
                pos_fin = c.get("pos_fin")
                if pos_ini and tam and pos_fin and pos_ini + tam - 1 != pos_fin:
                    tag = "erro"
                elif not pos_ini or not tam:
                    tag = "aviso"
                else:
                    tag = "par" if i % 2 == 0 else "impar"
                tree.insert("", tk.END, iid=str(i), tags=(tag,), values=vals)

        # Seleciona a aba padrão
        idx_padrao = nomes.index(aba_padrao) if aba_padrao in nomes else 0
        self._ignorar_tab_change = False
        self._nb_abas.select(idx_padrao)
        # Dispara manualmente (select pode não disparar evento se já estava naquele tab)
        self._on_tab_changed()

    def _on_tab_changed(self, _evt=None):
        """Chamado quando o usuário clica em outra aba do notebook."""
        if self._ignorar_tab_change or not self._nb_abas.tabs():
            return
        idx = self._nb_abas.index("current")
        nomes = list(self._dados_por_aba.keys())
        if idx >= len(nomes):
            return
        nome_aba = nomes[idx]
        self._aba_ativa = nome_aba
        aba = self._dados_por_aba.get(nome_aba, {})
        self._campos = aba.get("campos", [])
        self._headers_ativos = aba.get("headers", [])
        self._tree = self._trees_abas.get(nome_aba)
        self._lbl_count.config(text=f"{len(self._campos)} campos")
        self._atualizar_total()
        nome_xml = _nome_xml_para_aba(nome_aba)
        self._set_status(f"Aba: {nome_aba}  ({len(self._campos)} campos)  →  {nome_xml}")

    def _mudar_aba(self, nome_aba):
        """Seleciona a aba pelo nome (usado internamente)."""
        nomes = list(self._dados_por_aba.keys())
        if nome_aba in nomes:
            self._nb_abas.select(nomes.index(nome_aba))
            self._on_tab_changed()

    def _ler_xlsx_generico(self, path):
        """Lê a primeira aba como lista de campos, usando a 1ª linha como cabeçalho."""
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        headers = [_cell_str(c.value, f"Col{i}") for i, c in enumerate(ws[1])]
        campos = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            nome = _cell_str(row[0]) if row else ""
            if not nome:
                continue
            c = {"nome": nome, "entrada": "S", "tipo": "TEXTO"}
            for h, v in zip(headers[1:], row[1:]):
                c[_normalizar_chave(h)] = _cell_str(v)
            campos.append(c)
        return campos

    # ── Copiar campo da origem ────────────────────────────────────────────────

    def copiar_campo(self):
        if not self._dados_por_aba_origem:
            messagebox.showwarning("Aviso", "Carregue a planilha origem primeiro.")
            return

        def _processar(campos_selecionados):
            if not campos_selecionados:
                return

            novos, duplicatas = [], []
            for orig in campos_selecionados:
                nome = orig.get("nome", "")
                existente = next((c for c in self._campos if c.get("nome") == nome), None)
                if existente:
                    duplicatas.append((existente, orig))
                else:
                    novos.append(orig)

            # Adiciona campos novos sequencialmente
            for orig in novos:
                novo = dict(orig)
                ativos = [c for c in self._campos if c.get("pos_ini") and c.get("tamanho")]
                prox = (max(c["pos_ini"] + c["tamanho"] for c in ativos) if ativos else 1)
                novo["pos_ini"] = prox
                if novo.get("tamanho"):
                    novo["pos_fin"] = prox + novo["tamanho"] - 1
                self._campos.append(novo)

            # Pergunta sobre duplicatas de uma vez só
            atualizados = 0
            if duplicatas:
                nomes_dup = "\n".join(f"  • {e.get('nome')}" for e, _ in duplicatas)
                if messagebox.askyesno("Campos já existem",
                        f"Os seguintes campos já existem na planilha principal:\n{nomes_dup}\n\n"
                        "Deseja atualizar os atributos deles (tipo, tamanho, alinhamento etc.)?"):
                    for existente, orig in duplicatas:
                        for key in ("tipo", "tamanho", "alinhamento", "descricao",
                                    "obrigatorio", "coluna_db", "valor_padrao"):
                            if orig.get(key) is not None:
                                existente[key] = orig[key]
                        if not existente.get("valor"):
                            existente["valor"] = orig.get("valor_padrao", "")
                        atualizados += 1

            self._atualizar_tabela()
            partes = []
            if novos:
                partes.append(f"{len(novos)} copiado(s)")
            if atualizados:
                partes.append(f"{atualizados} atualizado(s)")
            if partes:
                self._set_status(f"Campos da origem: {', '.join(partes)}.")

        JanelaCopiarCampos(self.root, self._dados_por_aba_origem, on_confirmar=_processar)

    # ── Tabela ────────────────────────────────────────────────────────────────

    def _atualizar_tabela(self):
        """Repopula a Treeview da aba ativa com os dados atuais de self._campos."""
        if not self._tree:
            return

        tree    = self._tree
        headers = self._headers_ativos
        filtro  = self._var_filtro.get().lower()
        exib    = 0

        for item in tree.get_children():
            tree.delete(item)

        for i, c in enumerate(self._campos):
            nome = (c.get("nome") or "").lower()
            desc = (c.get("descricao") or "").lower()
            if filtro and filtro not in nome and filtro not in desc:
                continue

            raw  = c.get("_raw", {})
            vals = tuple(raw.get(h, "") for h in headers) if headers else (c.get("nome", ""),)

            pos_ini = c.get("pos_ini")
            tam     = c.get("tamanho")
            pos_fin = c.get("pos_fin")
            if pos_ini and tam and pos_fin and pos_ini + tam - 1 != pos_fin:
                tag = "erro"
            elif not pos_ini or not tam:
                tag = "aviso"
            else:
                tag = "par" if i % 2 == 0 else "impar"

            tree.insert("", tk.END, iid=str(i), tags=(tag,), values=vals)
            exib += 1

        self._lbl_count.config(text=f"{exib}/{len(self._campos)} campos")
        self._atualizar_total()

    def _atualizar_total(self):
        ativos = [c for c in self._campos if c.get("pos_ini") and c.get("tamanho")]
        if ativos:
            total = sum(c["tamanho"] for c in ativos)
            maxi  = max(c["pos_ini"] + c["tamanho"] - 1 for c in ativos)
            self._var_total.set(f"Tamanho total: {total} bytes | Pos. máx: {maxi}")
        else:
            self._var_total.set("")

    def _on_selecionar(self, _evt):
        pass  # pode expandir para pré-preencher detalhe

    def _on_duplo_clique(self, _evt):
        self.editar_campo()

    # ── Ações CRUD ────────────────────────────────────────────────────────────

    def novo_campo(self):
        # Sugere próxima posição
        ativos = [c for c in self._campos if c.get("pos_ini") and c.get("tamanho")]
        prox = (max(c["pos_ini"] + c["tamanho"] for c in ativos) if ativos else 1)
        JanelaEditarCampo(
            self.root,
            campo={"pos_ini": prox, "entrada": "S"},
            on_confirmar=self._adicionar_campo
        )

    def _adicionar_campo(self, campo):
        max_id = max(
            (int(c["id"]) for c in self._campos if str(c.get("id", "")).isdigit()),
            default=0
        )
        campo["id"] = str(max_id + 1)
        self._campos.append(campo)
        self._atualizar_tabela()
        self._set_status(f"Campo '{campo['nome']}' adicionado.")

    def editar_campo(self):
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("Aviso", "Selecione um campo na tabela.")
            return
        idx = int(sel[0])
        if idx >= len(self._campos):
            return

        def _aplicar(novo):
            novo["id"]    = self._campos[idx].get("id", "")
            novo["linha"] = self._campos[idx].get("linha", "")
            self._campos[idx] = novo
            self._atualizar_tabela()
            self._set_status(f"Campo '{novo['nome']}' atualizado.")

        JanelaEditarCampo(self.root, campo=self._campos[idx], on_confirmar=_aplicar)

    def remover_campo(self):
        sel = self._tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx >= len(self._campos):
            return
        nome = self._campos[idx].get("nome", "?")
        if messagebox.askyesno("Confirmar", f"Remover o campo '{nome}'?"):
            self._campos.pop(idx)
            self._atualizar_tabela()
            self._set_status(f"Campo '{nome}' removido.")

    def recalcular_posicoes(self):
        if not messagebox.askyesno("Recalcular Posições",
                "Recalcula posições de TODOS os campos de entrada em sequência a partir de 1.\n"
                "Continuar?"):
            return
        ativos = sorted(
            [c for c in self._campos if (c.get("entrada","S") or "S").upper()=="S" and c.get("tamanho")],
            key=lambda c: c.get("pos_ini", 99999)
        )
        pos = 1
        for c in ativos:
            c["pos_ini"] = pos
            c["pos_fin"] = pos + c["tamanho"] - 1
            pos += c["tamanho"]
        self._atualizar_tabela()
        self._set_status(f"Posições recalculadas. Total: {pos-1} bytes.")

    # ── Salvar planilha ───────────────────────────────────────────────────────

    def salvar_planilha(self):
        if not self._campos:
            messagebox.showwarning("Aviso", "Nenhum campo para salvar.")
            return

        path = self._arquivo_principal
        if not path:
            path = filedialog.asksaveasfilename(
                title="Salvar Planilha",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")]
            )
            if not path:
                return
            self._arquivo_principal = path

        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in (".xlsx", ".xls"):
                salvar_xlsx(path, self._campos, self._aba_ativa or "Campos Entrada")
            else:
                salvar_csv(path, self._campos)
            self._set_status(f"Planilha salva: {os.path.basename(path)}")
            messagebox.showinfo("Sucesso", f"Planilha salva:\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao salvar", str(e))

    # ── Validação ─────────────────────────────────────────────────────────────

    def validar(self):
        if not self._campos:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return

        erros, avisos, infos = validar_campos(self._campos)

        self._txt_val.configure(state=tk.NORMAL)
        self._txt_val.delete(1.0, tk.END)

        self._txt_val.insert(tk.END, "═" * 42 + "\n", "titulo")
        self._txt_val.insert(tk.END, "   RESULTADO DA VALIDAÇÃO\n", "titulo")
        self._txt_val.insert(tk.END, "═" * 42 + "\n\n", "titulo")

        if infos:
            self._txt_val.insert(tk.END, "INFORMAÇÕES:\n", "info")
            for i in infos:
                self._txt_val.insert(tk.END, f"  • {i}\n", "info")
            self._txt_val.insert(tk.END, "\n")

        if avisos:
            self._txt_val.insert(tk.END, "AVISOS:\n", "aviso")
            for a in avisos:
                self._txt_val.insert(tk.END, f"  ⚠  {a}\n", "aviso")
            self._txt_val.insert(tk.END, "\n")

        if erros:
            self._txt_val.insert(tk.END, "ERROS:\n", "erro")
            for e in erros:
                self._txt_val.insert(tk.END, f"  ✗  {e}\n", "erro")
            self._txt_val.insert(tk.END, "\n")
        else:
            self._txt_val.insert(tk.END, "✔  Sem erros de posicionamento!\n", "ok")

        self._txt_val.configure(state=tk.DISABLED)

        # Navega para aba de validação
        self._notebook.select(0)

        status = f"Validação: {'OK' if not erros else f'{len(erros)} erro(s)'}"
        if avisos:
            status += f" | {len(avisos)} aviso(s)"
        self._set_status(status)

    # ── XML ───────────────────────────────────────────────────────────────────

    def preview_xml(self):
        if not self._campos:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return
        try:
            xml_str = construir_xml(self._campos, self._headers_ativos)
            self._txt_xml.configure(state=tk.NORMAL)
            self._txt_xml.delete(1.0, tk.END)
            self._txt_xml.insert(tk.END, xml_str)
            self._txt_xml.configure(state=tk.DISABLED)
            self._notebook.select(1)
        except Exception as e:
            messagebox.showerror("Erro ao gerar XML", str(e))

    def gerar_xml(self):
        if not self._campos:
            messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
            return

        erros, _, _ = validar_campos(self._campos)
        if erros:
            if not messagebox.askyesno("Validação com erros",
                    f"Existem {len(erros)} erro(s) de validação.\nDeseja gerar o XML mesmo assim?"):
                return

        sugerido = ""
        if self._arquivo_principal:
            nome_xml = _nome_xml_para_aba(self._aba_ativa) if self._aba_ativa else "XML"
            dir_base = os.path.dirname(self._arquivo_principal)
            sugerido = os.path.join(dir_base, nome_xml + ".xml")

        path = filedialog.asksaveasfilename(
            title="Salvar XML",
            defaultextension=".xml",
            filetypes=[("XML", "*.xml"), ("Todos", "*.*")],
            initialfile=os.path.basename(sugerido) if sugerido else "evento.xml",
            initialdir=os.path.dirname(sugerido) if sugerido else ""
        )
        if not path:
            return

        try:
            xml_str = construir_xml(self._campos, self._headers_ativos)
            with open(path, "w", encoding="utf-8") as f:
                f.write(xml_str)

            # Atualiza preview
            self._txt_xml.configure(state=tk.NORMAL)
            self._txt_xml.delete(1.0, tk.END)
            self._txt_xml.insert(tk.END, xml_str)
            self._txt_xml.configure(state=tk.DISABLED)
            self._notebook.select(1)

            ativos = [c for c in self._campos
                      if (c.get("entrada","S") or "S").upper()=="S" and c.get("pos_ini")]
            total = sum(c.get("tamanho", 0) or 0 for c in ativos)
            self._set_status(
                f"XML gerado: {os.path.basename(path)}  |  "
                f"{len(ativos)} campos  |  {total} bytes"
            )
            messagebox.showinfo("Sucesso", f"XML gerado com sucesso!\n{path}")
        except Exception as e:
            messagebox.showerror("Erro ao gerar XML", str(e))

    # ── Utilidades ────────────────────────────────────────────────────────────

    def _set_status(self, msg):
        self._var_status.set(msg)


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.25)
    except Exception:
        pass
    GeradorXMLApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
