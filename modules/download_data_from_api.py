import requests
from colorama import init, Fore, Style

init(autoreset=True)


def download_data_from_api(
    url: str, json_file_path: str, chunk_size: int = 1000,
    message: str = 'Загружено байтов: '
):
    """Получаем данные из API и записываем их в json файл."""
    response = requests.get(url, stream=True)

    if response.status_code == 200:
        loaded_size = 0

        with open(json_file_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=chunk_size):
                if chunk:
                    loaded_size += len(chunk)
                    f.write(chunk)

                    print(
                        Fore.WHITE + Style.DIM + message +
                        str(loaded_size), end='\r'
                    )
            print()

    else:
        print(f'{Fore.RED}Ошибка при запросе. Проверьте ссылку: {url}')
