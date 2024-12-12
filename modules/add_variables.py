import os
import shutil

import pandas as pd
from typing import Optional
from docxtpl import DocxTemplate

from modules.add_table_variables import add_table_from_dataframe


def add_variables(
    context: dict,
    template_path: str,
    output_name: str,
    output_dir: str,
    table: Optional[pd.DataFrame] = None
):
    """
    Создание документа на основе шаблона с заменой переменных и добавлением
    таблицы.
    """
    output_path = os.path.join(output_dir, output_name)

    # Создание копии шаблона:
    shutil.copy(template_path, output_path)

    # Загрузка шаблона Word:
    doc = DocxTemplate(output_path)

    # Рендеринг документа с заменой переменных:
    doc.render(context)

    # Если таблица предоставлена, добавляем ее сразу после рендеринга:
    if table is not None:
        add_table_from_dataframe(doc, table)

    # Сохранение финального документа:
    doc.save(output_path)
