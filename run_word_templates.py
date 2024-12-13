import os
import sys
from datetime import datetime
from typing import Optional
from concurrent.futures import ThreadPoolExecutor

import pandas as pd
from colorama import Fore, Style, init

from modules.api_links import api_links
from modules.tmp_files_names import tmp_files_names
from modules.download_data_from_api import download_data_from_api
from modules.prepare_folders import prepare_folders
from modules.prepare_output_folders import prepare_output_folders
from modules.write_results import write_results
from modules.group_poles import group_poles
from modules.function_execution_time import function_execution_time
from modules.make_templates import make_templates


init(autoreset=True)

# Проверяем, запущен ли скрипт как exe:
if getattr(sys, 'frozen', False):
    CURRENT_DIR = os.path.join(os.path.dirname(sys.executable), '..')
else:
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

INPUT_FOLDER_DIR: str = os.path.join(CURRENT_DIR, 'input')
TEMPLATES_FOLDER_DIR: str = os.path.join(CURRENT_DIR, 'templates')
TMP_FOLDER_DIR: str = os.path.join(CURRENT_DIR, 'tmp')
OUTPUT_FOLDER_DIR: str = os.path.join(CURRENT_DIR, 'output')
LOG_FOLDER_DIR: str = os.path.join(CURRENT_DIR, 'log')
OUTPUT_LOG_PATH: str = os.path.join(LOG_FOLDER_DIR, 'word_templates.csv')
NORMAL_INPUT_TEMPLATE_PATH: str = os.path.join(
    TEMPLATES_FOLDER_DIR, 'normal.docx'
)
NO_TENANT_INPUT_TEMPLATE_PATH: str = os.path.join(
    TEMPLATES_FOLDER_DIR, 'no_tenant.docx'
)

POLE_PART: int = 8   # Первые n символов шифра опоры для поиска.
MAX_WORKERS = os.cpu_count()

# Объединяем данные отчетов по оборудованию и kits по AssetKitId, чтобы
# определить шифр опоры, а затем уже объеденяем с договарами по шифру опоры.
COLUMNS_TO_KEEP_OPERATORS: list[str] = [
    'AssetKitId',
    'Оператор',
    'Тип',
    'Верхняя отм. оборудования',
    'Ширина',
    'Вес',
    'Диаметр',
    'Азимут'
]
COLUMNS_TO_KEEP_KITS: list[str] = ['AssetKitId', 'Шифр опоры']
COLUMNS_TO_KEEP_CONTRACTS: list[str] = [
    'Шифр опоры', 'Номер рамочного договора'
]
COLUMNS_TO_KEEP_TL: list[str] = ['Шифр', 'Статус опоры']
COLUMNS_TO_KEEP_JOIN: list[str] = [
    'Оператор',
    'Номер рамочного договора',
    'Тип',
    'Верхняя отм. оборудования',
    'Ширина',
    'Вес',
    'Диаметр',
    'Азимут'
]


@function_execution_time
def prepare_word_templates(
    user_name: str, user_email: str, user_phone: str, update_data: bool = True,
    notification_date: Optional[str] = None
):
    """
    Генерация шаблонов word 'Уведомления РБТ'.

    Parameters
    ----------
    user_name: str
        ФИО исполнителя.
    user_email: str
        Email исполнителя.
    user_phone: str
        Моб. тел. исполнителя.
    update_data: bool = True
        Обновить данные из TowerStore.
    """

    notification_date = notification_date or datetime.now().strftime(
        '%d.%m.%Y'
    )

    prepare_folders(
        [
            INPUT_FOLDER_DIR,
            TEMPLATES_FOLDER_DIR,
            TMP_FOLDER_DIR,
            OUTPUT_FOLDER_DIR,
            LOG_FOLDER_DIR
        ]
    )

    prepare_output_folders([OUTPUT_FOLDER_DIR])

    if update_data:
        download_data_from_api(
            api_links.KITS,
            os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.KITS),
            message='KITS -- загрузка байтов: '
        )
        download_data_from_api(
            api_links.TL,
            os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.TL),
            message='TL -- загрузка байтов: '
        )
        download_data_from_api(
            api_links.CONTRACT,
            os.path.join(
                CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.CONTRACT
            ),
            message='Рамочные договора с операторами -- загрузка байтов: '
        )
        download_data_from_api(
            api_links.MTS,
            os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.MTS),
            message='МТС -- загрузка байтов: '
        )
        download_data_from_api(
            api_links.MEGAFON,
            os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.MEGAFON),
            message='Мегафон -- загрузка байтов: '
        )
        download_data_from_api(
            api_links.TELE_2,
            os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.TELE_2),
            message='Теле2 -- загрузка байтов: '
        )
        download_data_from_api(
            api_links.VIMPELCOM,
            os.path.join(
                CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.VIMPELCOM
            ),
            message='Вымпелком -- загрузка байтов: '
        )
        download_data_from_api(
            api_links.OTHER,
            os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.OTHER),
            message='Прочее -- загрузка байтов: '
        )

    print(Fore.BLACK + Style.DIM + 'Пожалуйста подождите...')

    base_df = pd.read_excel(
        os.path.join(CURRENT_DIR, INPUT_FOLDER_DIR, 'Список АМС РБТ.xlsx')
    )

    valid_poles: tuple[str] = tuple(set(
        [
            str(pole).strip()[:POLE_PART]
            for pole in base_df['Шифр TS'].to_list()
        ]
    ))

    main_df: pd.DataFrame = pd.concat(
        [
            pd.read_json(
                os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.MTS),
                orient='records'
            )[COLUMNS_TO_KEEP_OPERATORS],
            pd.read_json(
                os.path.join(
                    CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.MEGAFON
                ),
                orient='records'
            )[COLUMNS_TO_KEEP_OPERATORS],
            pd.read_json(
                os.path.join(
                    CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.TELE_2
                ),
                orient='records'
            )[COLUMNS_TO_KEEP_OPERATORS],
            pd.read_json(
                os.path.join(
                    CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.VIMPELCOM
                ),
                orient='records'
            )[COLUMNS_TO_KEEP_OPERATORS],
            pd.read_json(
                os.path.join(
                    CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.OTHER
                ),
                orient='records'
            )[COLUMNS_TO_KEEP_OPERATORS],
        ],
        ignore_index=True
    )

    main_df = pd.merge(
        main_df,
        pd.read_json(
            os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.KITS),
            orient='records'
        )[COLUMNS_TO_KEEP_KITS],
        on='AssetKitId', how='left'
    )

    main_df = pd.merge(
        main_df,
        pd.read_json(
            os.path.join(
                CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.CONTRACT
            ),
            orient='records'
        )[COLUMNS_TO_KEEP_CONTRACTS],
        on='Шифр опоры', how='left'
    )

    filter_columns = (
        COLUMNS_TO_KEEP_JOIN + ['Шифр опоры']
    ) if 'Шифр опоры' not in COLUMNS_TO_KEEP_JOIN else COLUMNS_TO_KEEP_JOIN

    main_df = main_df[filter_columns]

    main_df = main_df[
        main_df['Шифр опоры'].str.startswith(valid_poles, na=False)
    ].reset_index(drop=True)

    tl_df = pd.read_json(
        os.path.join(CURRENT_DIR, TMP_FOLDER_DIR, tmp_files_names.TL),
        orient='records'
    )[COLUMNS_TO_KEEP_TL]

    tl_df = tl_df[
        tl_df['Шифр'].str.startswith(valid_poles, na=False)
    ].reset_index(drop=True)

    COUNT_BASE_DF: int = len(base_df)

    base_df = base_df.fillna('н/д')
    main_df = main_df.fillna('н/д')
    tl_df = tl_df.fillna('н/д')

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {
            executor.submit(
                make_templates,
                row, index,
                POLE_PART, COUNT_BASE_DF, COLUMNS_TO_KEEP_JOIN,
                tl_df, main_df,
                NORMAL_INPUT_TEMPLATE_PATH, NO_TENANT_INPUT_TEMPLATE_PATH,
                OUTPUT_LOG_PATH, OUTPUT_FOLDER_DIR,
                notification_date, user_name, user_email, user_phone,
            ): index for index, row in base_df.iterrows()
        }
        for future in futures:
            future.result()

    group_poles(OUTPUT_FOLDER_DIR, base_df)


if __name__ == '__main__':
    user_name = input(Fore.CYAN + 'ФИО исполнителя: ' + Style.RESET_ALL)
    user_email = input(Fore.CYAN + 'Email исполнителя: ' + Style.RESET_ALL)
    user_phone = input(Fore.CYAN + 'Тел. исполнителя: ' + Style.RESET_ALL)

    update_data = input(
        f'{Style.BRIGHT}Обновить данные (Y/N): ' + Style.RESET_ALL
    ).lower()

    if update_data == 'y':
        update_data = True
    elif update_data == 'n':
        update_data = False
    else:
        print(Fore.RED + 'Данные введены некорректно, повторите попытку.')
        exit(1)

    current_date_time: str = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
    log_message: str = (
        f'TIMESTAMP: {current_date_time}, '
        + f'USER: {user_name}, EMAIL: {user_email}, '
        + f'PHONE: {user_phone}, UPDATE_DATA: {update_data}'
    )
    write_results(OUTPUT_LOG_PATH, log_message)
    prepare_word_templates(user_name, user_email, user_phone, update_data)
    write_results(OUTPUT_LOG_PATH, '\n\n')
    while True:
        exit_check = input(
            Fore.WHITE + Style.BRIGHT +
            'Для выхода введите Q: ' + Style.RESET_ALL
        )
        if exit_check.strip().lower() == 'q':
            break


# 60502-30-30-0-2 добавь н/д если нет номер рамочного договора.
