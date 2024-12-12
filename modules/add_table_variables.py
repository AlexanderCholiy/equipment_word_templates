import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_table_from_dataframe(
        doc: Document, df: pd.DataFrame, font_size: int = 9
):
    """Добавление таблицы из DataFrame в документ с форматированием."""

    # Создание таблицы:
    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])

    # Настройка стиля таблицы:
    tbl = table._element
    tblBorders = parse_xml(
        r'<w:tblBorders %s>'
        r'<w:top w:val="single" w:sz="4" w:space="0"/>'
        r'<w:left w:val="single" w:sz="4" w:space="0"/>'
        r'<w:bottom w:val="single" w:sz="4" w:space="0"/>'
        r'<w:right w:val="single" w:sz="4" w:space="0"/>'
        r'<w:insideH w:val="single" w:sz="4" w:space="0"/>'
        r'<w:insideV w:val="single" w:sz="4" w:space="0"/>'
        r'</w:tblBorders>' % nsdecls('w')
    )
    tblPr = tbl.find(
        './/w:tblPr',
        namespaces={
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
    )
    if tblPr is None:
        tblPr = parse_xml(r'<w:tblPr %s/>' % nsdecls('w'))
        tbl.append(tblPr)
    tblPr.append(tblBorders)

    # Добавление заголовков:
    header_cells = table.rows[0].cells
    for i, column_name in enumerate(df.columns):
        cell = header_cells[i]
        cell.text = column_name
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.runs[0]
            run.bold = True
            run.font.size = Pt(font_size)
            run.font.name = 'Times New Roman'

    # Добавление данных:
    for i, row in df.iterrows():
        for j, value in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(value)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.runs[0]
                run.font.size = Pt(font_size)
                run.font.name = 'Times New Roman'

    # Центрирование таблицы в документе:
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Вставка таблицы в нужное место:
    for paragraph in doc.paragraphs:
        if 'var_table' in paragraph.text:
            new_paragraph = paragraph.insert_paragraph_before('')
            new_paragraph._element.addnext(table._element)
            p_element = paragraph._element
            p_element.getparent().remove(p_element)
            break
