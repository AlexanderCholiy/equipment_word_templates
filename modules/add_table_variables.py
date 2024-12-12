import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_table_from_dataframe(
    doc: Document, df: pd.DataFrame, font_size: int = 9
) -> None:
    """Добавление таблицы из DataFrame в документ с форматированием."""

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

    # Добавление заголовков с настройкой размера шрифта:
    for i, column_name in enumerate(df.columns):
        cell = table.cell(0, i)
        run = cell.paragraphs[0].add_run(column_name)
        run.bold = True
        run.font.size = Pt(font_size)
        run.font.name = 'Times New Roman'
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Добавление данных с настройкой размера шрифта:
    for i, row in df.iterrows():
        for j, value in enumerate(row):
            cell = table.cell(i + 1, j)
            run = cell.paragraphs[0].add_run(str(value))
            run.font.size = Pt(font_size)
            run.font.name = 'Times New Roman'
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

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
