"""
core/spreadsheet.py — Manipulação de planilhas Excel para o Arqui Specs.

Modelo definido com a Gabriela em 2026-05-25:
- Cada ambiente fica em uma aba própria do arquivo Excel.
- A aba "Info" guarda metadados do projeto.
- Campos de cada aba: Item, Modelo, Segmento, Acabamento, Marca, Dimensões, Quantidade.

O app mostra um preview agregado com a coluna AMBIENTE, mas o arquivo baixado
fica separado por abas para ficar mais natural para uso profissional.
"""

from io import BytesIO
import datetime
import re

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd


FIELDS = [
    ("item", "ITEM", 24),
    ("modelo", "MODELO", 24),
    ("segmento", "SEGMENTO", 22),
    ("acabamento", "ACABAMENTO", 22),
    ("marca", "MARCA", 18),
    ("dimensoes", "DIMENSÕES", 24),
    ("quantidade", "QUANTIDADE", 14),
]

ID_COL = len(FIELDS) + 1
HEADER_ROWS = 2
DATA_START_ROW = HEADER_ROWS + 1
SYSTEM_SHEETS = {"Info"}


class Colors:
    BG_PRIMARY = "3D1812"      # Bordô escuro
    ACCENT = "7A2E2A"          # Terracota
    SAND = "D9CBA8"            # Bege areia
    CREAM = "F4ECDC"           # Off-white
    INK = "1F1410"             # Marrom quase preto
    BORDER = "D9CBA8"

    PROJECT_HEADER_BG = BG_PRIMARY
    PROJECT_HEADER_FONT = CREAM
    COL_HEADER_BG = ACCENT
    COL_HEADER_FONT = CREAM
    ROW_ODD = CREAM
    ROW_EVEN = "E8DDBF"


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def _font(bold=False, color=Colors.INK, size=10) -> Font:
    return Font(bold=bold, color=color, size=size, name="Calibri")


def _border_thin() -> Border:
    thin = Side(style="thin", color=Colors.BORDER)
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _align(horizontal="left", wrap=True) -> Alignment:
    return Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap)


def sanitize_sheet_name(name: str) -> str:
    """Converte um nome de ambiente em nome válido de aba Excel."""
    cleaned = re.sub(r"[\[\]\*:/\\?]", "-", (name or "AMBIENTE").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned[:31] or "AMBIENTE"


def _is_environment_sheet(ws) -> bool:
    return ws.title not in SYSTEM_SHEETS


def _setup_environment_sheet(ws, ambiente: str) -> None:
    """Aplica estrutura e estilos padrão em uma aba de ambiente."""
    ws.title = sanitize_sheet_name(ambiente)
    ws.sheet_view.showGridLines = False

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(FIELDS))
    ws["A1"] = ambiente.upper()
    ws["A1"].font = _font(bold=True, color=Colors.PROJECT_HEADER_FONT, size=13)
    ws["A1"].fill = _fill(Colors.PROJECT_HEADER_BG)
    ws["A1"].alignment = _align(horizontal="center", wrap=False)
    ws.row_dimensions[1].height = 32

    for idx, (_, label, width) in enumerate(FIELDS, start=1):
        cell = ws.cell(row=2, column=idx)
        cell.value = label
        cell.font = _font(bold=True, color=Colors.COL_HEADER_FONT, size=10)
        cell.fill = _fill(Colors.COL_HEADER_BG)
        cell.alignment = _align(horizontal="center", wrap=False)
        cell.border = _border_thin()
        ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = width

    ws.column_dimensions[openpyxl.utils.get_column_letter(ID_COL)].hidden = True
    ws.freeze_panes = "A3"


def create_template(
    project_name: str = "Memorial Descritivo",
    client_name: str = "(definir)",
    responsible: str = "Gabriela Lendecker · CAU A123456-7",
) -> BytesIO:
    """Cria workbook inicial com uma aba modelo e a aba Info."""
    wb = openpyxl.Workbook()
    ws = wb.active
    _setup_environment_sheet(ws, "Ambiente modelo")

    ws_info = wb.create_sheet("Info")
    ws_info["A1"] = "Nome do Projeto"
    ws_info["B1"] = project_name
    ws_info["A2"] = "Cliente"
    ws_info["B2"] = client_name
    ws_info["A3"] = "Responsável"
    ws_info["B3"] = responsible
    ws_info["A4"] = "Emissão"
    ws_info["B4"] = datetime.date.today().strftime("%d %b %Y")
    ws_info["A5"] = "Versão"
    ws_info["B5"] = "v1"
    ws_info.column_dimensions["A"].width = 20
    ws_info.column_dimensions["B"].width = 45
    for row in range(1, 6):
        label = ws_info.cell(row=row, column=1)
        value = ws_info.cell(row=row, column=2)
        label.fill = _fill(Colors.ACCENT)
        label.font = _font(bold=True, color=Colors.CREAM)
        label.alignment = _align(horizontal="left", wrap=False)
        label.border = _border_thin()
        value.fill = _fill(Colors.CREAM)
        value.font = _font(color=Colors.INK)
        value.alignment = _align(horizontal="left", wrap=False)
        value.border = _border_thin()

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def load_workbook_from_bytes(data: bytes) -> openpyxl.Workbook:
    return openpyxl.load_workbook(BytesIO(data))


def save_workbook_to_bytes(wb: openpyxl.Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _next_id(wb: openpyxl.Workbook) -> int:
    max_id = 0
    for ws in wb.worksheets:
        if not _is_environment_sheet(ws):
            continue
        for row in ws.iter_rows(min_row=DATA_START_ROW, min_col=ID_COL, max_col=ID_COL, values_only=True):
            value = row[0]
            if isinstance(value, int):
                max_id = max(max_id, value)
    return max_id + 1


def _get_or_create_environment_sheet(wb: openpyxl.Workbook, ambiente: str):
    sheet_name = sanitize_sheet_name(ambiente.upper())
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.create_sheet(sheet_name)
        _setup_environment_sheet(ws, ambiente.upper())

    if "Ambiente modelo" in wb.sheetnames and sheet_name != "Ambiente modelo":
        model_ws = wb["Ambiente modelo"]
        if len(get_rows_from_sheet(model_ws, "Ambiente modelo")) == 0:
            del wb["Ambiente modelo"]

    return ws


def _next_data_row(ws) -> int:
    row = max(ws.max_row + 1, DATA_START_ROW)
    while True:
        if not any(ws.cell(row=row, column=col).value for col in range(1, len(FIELDS) + 1)):
            return row
        row += 1


def _apply_data_row_style(ws, row_num: int) -> None:
    data_index = row_num - DATA_START_ROW
    bg = Colors.ROW_ODD if data_index % 2 == 0 else Colors.ROW_EVEN
    for col_idx in range(1, len(FIELDS) + 1):
        cell = ws.cell(row=row_num, column=col_idx)
        cell.fill = _fill(bg)
        cell.font = _font()
        cell.alignment = _align()
        cell.border = _border_thin()
    ws.row_dimensions[row_num].height = 32


def get_rows_from_sheet(ws, ambiente: str) -> list[dict]:
    rows = []
    for row_num in range(DATA_START_ROW, ws.max_row + 1):
        values = [ws.cell(row=row_num, column=i).value for i in range(1, len(FIELDS) + 1)]
        if not any(values):
            continue
        row_id = ws.cell(row=row_num, column=ID_COL).value
        item = {"id": row_id, "ambiente": ambiente}
        for idx, (key, _, _) in enumerate(FIELDS):
            item[key] = values[idx] or ""
        rows.append(item)
    return rows


def get_sheet_data(wb: openpyxl.Workbook) -> list[dict]:
    rows = []
    for ws in wb.worksheets:
        if _is_environment_sheet(ws):
            rows.extend(get_rows_from_sheet(ws, ws.title.upper()))
    return rows


def get_compact_context(wb: openpyxl.Workbook) -> str:
    data = get_sheet_data(wb)
    if not data:
        return "📋 Planilha vazia — pronta para receber itens. Estrutura: uma aba por ambiente."

    by_amb: dict[str, list[str]] = {}
    for item in data:
        amb = item["ambiente"] or "SEM AMBIENTE"
        by_amb.setdefault(amb, []).append(str(item.get("item") or "(sem item)"))

    total = len(data)
    lines = [f"📋 Planilha atual: {total} {'item' if total == 1 else 'itens'} em {len(by_amb)} {'ambiente' if len(by_amb) == 1 else 'ambientes'}"]
    for amb, items in by_amb.items():
        lines.append(f"  • Aba {amb}: {', '.join(items[:8])}")
    return "\n".join(lines)


def get_project_info(wb: openpyxl.Workbook) -> dict:
    try:
        ws = wb["Info"]
        return {
            "nome": ws["B1"].value or "Memorial Descritivo",
            "cliente": ws["B2"].value or "(definir)",
            "responsavel": ws["B3"].value or "Gabriela Lendecker",
            "emissao": ws["B4"].value or "",
            "versao": ws["B5"].value or "v1",
        }
    except Exception:
        return {"nome": "Memorial Descritivo", "cliente": "", "responsavel": "", "emissao": "", "versao": "v1"}


def update_project_info(wb: openpyxl.Workbook, **kwargs) -> None:
    try:
        ws_info = wb["Info"]
    except KeyError:
        ws_info = wb.create_sheet("Info")

    mapping = {
        "nome": "B1",
        "cliente": "B2",
        "responsavel": "B3",
        "versao": "B5",
    }
    for field, cell in mapping.items():
        if field in kwargs:
            ws_info[cell] = kwargs[field]
    ws_info["B4"] = datetime.date.today().strftime("%d %b %Y")


def _update_version(wb: openpyxl.Workbook) -> None:
    try:
        ws_info = wb["Info"]
        version_str = ws_info["B5"].value or "v1"
        num = int(str(version_str).replace("v", "").replace("r", "").strip())
        ws_info["B5"] = f"v{num + 1}"
    except Exception:
        pass


def _find_row_by_id(wb: openpyxl.Workbook, linha_id: int):
    for ws in wb.worksheets:
        if not _is_environment_sheet(ws):
            continue
        for row_num in range(DATA_START_ROW, ws.max_row + 1):
            if ws.cell(row=row_num, column=ID_COL).value == linha_id:
                return ws, row_num
    return None, None


def apply_operations(wb: openpyxl.Workbook, operacoes: list[dict]) -> dict:
    adicionados = []
    atualizados = []
    erros = []

    for op in operacoes:
        acao = op.get("acao")
        if acao == "adicionar_item":
            try:
                ambiente = (op.get("ambiente") or "").strip().upper()
                if not ambiente:
                    erros.append("adicionar_item sem ambiente")
                    continue

                ws = _get_or_create_environment_sheet(wb, ambiente)
                row_num = _next_data_row(ws)
                row_id = _next_id(wb)

                for col_idx, (key, _, _) in enumerate(FIELDS, start=1):
                    ws.cell(row=row_num, column=col_idx).value = op.get(key, "")
                ws.cell(row=row_num, column=ID_COL).value = row_id
                _apply_data_row_style(ws, row_num)
                adicionados.append(f"{ambiente} · {op.get('item', '?')}")
            except Exception as e:
                erros.append(f"adicionar_item erro: {str(e)}")

        elif acao == "atualizar_item":
            linha_id = op.get("linha_id")
            if not linha_id:
                erros.append("atualizar_item sem linha_id")
                continue

            ws, row_num = _find_row_by_id(wb, linha_id)
            if ws is None:
                erros.append(f"linha ID {linha_id} não encontrada")
                continue

            novo_ambiente = op.get("ambiente")
            if novo_ambiente and sanitize_sheet_name(novo_ambiente.upper()) != ws.title:
                old_values = {
                    key: ws.cell(row=row_num, column=idx).value
                    for idx, (key, _, _) in enumerate(FIELDS, start=1)
                }
                old_values.update({k: v for k, v in op.items() if k in {f[0] for f in FIELDS}})
                ws.delete_rows(row_num)
                new_ws = _get_or_create_environment_sheet(wb, novo_ambiente.upper())
                new_row = _next_data_row(new_ws)
                for col_idx, (key, _, _) in enumerate(FIELDS, start=1):
                    new_ws.cell(row=new_row, column=col_idx).value = old_values.get(key, "")
                new_ws.cell(row=new_row, column=ID_COL).value = linha_id
                _apply_data_row_style(new_ws, new_row)
                atualizados.append(f"ID {linha_id} · movido para {novo_ambiente.upper()}")
            else:
                for col_idx, (key, _, _) in enumerate(FIELDS, start=1):
                    if key in op and op[key] is not None:
                        ws.cell(row=row_num, column=col_idx).value = op[key]
                _apply_data_row_style(ws, row_num)
                atualizados.append(f"ID {linha_id} · {op.get('item', '?')}")
        else:
            erros.append(f"ação desconhecida: {acao}")

    try:
        wb["Info"]["B4"] = datetime.date.today().strftime("%d %b %Y")
        _update_version(wb)
    except Exception:
        pass

    return {
        "items_adicionados": len(adicionados),
        "items_atualizados": len(atualizados),
        "adicionados": adicionados,
        "atualizados": atualizados,
        "erros": erros,
        "total_itens": len(get_sheet_data(wb)),
    }


def get_preview_dataframe(wb: openpyxl.Workbook) -> pd.DataFrame:
    data = get_sheet_data(wb)
    columns = ["AMBIENTE", "ITEM", "MODELO", "SEGMENTO", "ACABAMENTO", "MARCA", "DIMENSÕES", "QUANTIDADE"]
    if not data:
        return pd.DataFrame(columns=columns)

    df = pd.DataFrame(data)
    df = df.rename(columns={
        "ambiente": "AMBIENTE",
        "item": "ITEM",
        "modelo": "MODELO",
        "segmento": "SEGMENTO",
        "acabamento": "ACABAMENTO",
        "marca": "MARCA",
        "dimensoes": "DIMENSÕES",
        "quantidade": "QUANTIDADE",
    })
    if "id" in df.columns:
        df = df.drop(columns=["id"])
    return df[columns]


def style_preview_dataframe(df: pd.DataFrame) -> pd.io.formats.style.Styler:
    """Aplica estilo mínimo ao dataframe de preview.
    Mantém texto legível e header com identidade visual;
    deixa o fundo das linhas para o tema nativo do Streamlit.
    """
    styler = df.style.set_properties(**{
        "font-size": "12px",
        "color": "#2A1510",
    })
    styler = styler.set_table_styles([
        {"selector": "th", "props": [
            ("background-color", "#7A2E2A"),
            ("color", "#F4ECDC"),
            ("font-weight", "600"),
            ("font-size", "12px"),
        ]},
    ])
    return styler
