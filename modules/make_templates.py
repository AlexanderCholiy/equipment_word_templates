import threading

import pandas as pd
from pandas import Series
from colorama import Fore, Style, init

from modules.add_variables import add_variables
from modules.write_results import write_results


init(autoreset=True)

print_lock = threading.Lock()


def make_templates(
    row: Series, index: int,
    POLE_PART: int, COUNT_BASE_DF: int, COLUMNS_TO_KEEP_JOIN: list[str],
    tl_df: pd.DataFrame, main_df: pd.DataFrame,
    NORMAL_INPUT_TEMPLATE_PATH: str, NO_TENANT_INPUT_TEMPLATE_PATH: str,
    OUTPUT_LOG_PATH: str, OUTPUT_FOLDER_DIR: str,
    notification_date: str, user_name: str, user_email: str, user_phone: str
):
    try:
        no_tenant_flag = False
        table = None
        rcs = str(row['РЦС']).strip()
        pole = str(row['Шифр TS']).strip()
        rzd_ams_lease_agreement_number = str(
            row['№ договора аренды АМС "РЖД"']
        ).strip()
        date_of_contract = row['дата договора']
        if isinstance(date_of_contract, pd.Timestamp):
            date_of_contract = date_of_contract.strftime('%d-%m-%Y')
        else:
            date_of_contract = str(date_of_contract).strip()

        ams_name = str(row['Наименование']).strip()
        ams_address = str(row['Адрес']).strip()
        ams_id = str(row['ID АМС']).strip()
        ams_cadastral_number = str(row['Кадастровый номер участка']).strip()
        ams_square = str(row['площадь участка']).strip()
        ams_height = str(row['высота, м']).strip()

        output_name = f'{pole}.docx'

        filtered_df_tl = tl_df[
            tl_df['Шифр'].str.startswith(pole[:POLE_PART])
        ].reset_index(drop=True)

        if filtered_df_tl.empty:
            with print_lock:
                print(
                    Fore.RED +
                    f'{index+1}/{COUNT_BASE_DF} ' +
                    f'-- опора {pole} не найдена в TL;'
                )
                write_results(
                    OUTPUT_LOG_PATH,
                    f'{index+1}/{COUNT_BASE_DF} ' +
                    f'-- опора {pole} не найдена в TL;'
                )
            return

        if (filtered_df_tl['Статус опоры'] == 'No Tenant').any():
            no_tenant_flag = True

        if not no_tenant_flag:
            filtered_main_df = main_df[
                main_df['Шифр опоры'].str.startswith(pole[:POLE_PART])
            ].reset_index(drop=True)

            if not filtered_main_df.empty:
                table = filtered_main_df[COLUMNS_TO_KEEP_JOIN]

        template_path = NORMAL_INPUT_TEMPLATE_PATH if table is not None else (
            NO_TENANT_INPUT_TEMPLATE_PATH
        )

        add_variables(
            context={
                'rzd_ams_lease_agreement_number':
                rzd_ams_lease_agreement_number,
                'date_of_contract': date_of_contract,
                'ams_name': ams_name,
                'ams_address': ams_address,
                'ams_id': ams_id,
                'ams_cadastral_number': ams_cadastral_number,
                'ams_square': ams_square,
                'ams_height': ams_height,
                'notification_date': notification_date,
                'user_name': user_name,
                'user_email': user_email,
                'user_phone': user_phone,
            },
            template_path=template_path,
            output_name=output_name,
            output_dir=OUTPUT_FOLDER_DIR,
            table=table,
        )

        with print_lock:
            if no_tenant_flag:
                print(
                    Fore.LIGHTGREEN_EX + Style.DIM +
                    f'{index+1}/{COUNT_BASE_DF} ' +
                    '-- шаблон для {pole} ({rcs}) (No Tenant) готов;'
                )
                write_results(
                    OUTPUT_LOG_PATH,
                    f'{index+1}/{COUNT_BASE_DF} ' +
                    '-- шаблон для {pole} ({rcs}) (No Tenant) готов;'
                )
            elif table is None:
                print(
                    Fore.LIGHTYELLOW_EX + Style.DIM +
                    f'{index+1}/{COUNT_BASE_DF} ' +
                    f'-- шаблон для {pole} ({rcs}) ' +
                    '(отсутствует в отчёте по оборудованию) готов;'
                )
                write_results(
                    OUTPUT_LOG_PATH,
                    f'{index+1}/{COUNT_BASE_DF} ' +
                    f'-- шаблон для {pole} ({rcs}) ' +
                    '(отсутствует в отчёте по оборудованию) готов;'
                )
            else:
                print(
                    Fore.LIGHTGREEN_EX + Style.DIM +
                    f'{index+1}/{COUNT_BASE_DF} '
                    f'-- шаблон для {pole} ({rcs}) готов;'
                )
                write_results(
                    OUTPUT_LOG_PATH,
                    f'{index+1}/{COUNT_BASE_DF} ' +
                    f'-- шаблон для {pole} ({rcs}) готов;'
                )

    except Exception as e:
        with print_lock:
            print(Fore.RED + str(e))
            write_results(OUTPUT_LOG_PATH, str(e))
