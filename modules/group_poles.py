import os
import shutil

import pandas as pd


def clean_folder_name(folder_name: str) -> str:
    """Очистка имени папки от недопустимых символов."""
    return ''.join(
        c if c.isalnum() or c in (' ', '_') else '' for c in folder_name
    )


def group_poles(output_dir: str, base_notification_df: pd.DataFrame):
    """Группировка шаблонов по РЦС"""
    files_name: list[str] = os.listdir(output_dir)
    poles: tuple[str] = tuple(
        (file_name.replace('.docx', '') for file_name in files_name)
    )

    base_notification_df['Шифр TS'] = (
        base_notification_df['Шифр TS'].fillna('')
    )

    for index, file_name in enumerate(files_name):
        filtered_df = (
            base_notification_df[base_notification_df[
                'Шифр TS'
            ].str.startswith(poles[index][:8])].reset_index(drop=True)
        )
        if filtered_df.empty:
            continue
        rcs = filtered_df.iloc[0]["РЦС"]
        current_folder: str = os.path.join(output_dir, clean_folder_name(rcs))
        current_file_path: str = os.path.join(output_dir, file_name)
        if not os.path.exists(current_folder):
            os.makedirs(current_folder)
        shutil.move(current_file_path, current_folder)
