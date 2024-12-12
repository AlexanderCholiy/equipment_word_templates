def write_results(file_path: str, message: str) -> None:
    """Запись результатов в .csv файл."""
    with open(file_path, mode='a', encoding='UTF-8') as file:
        file.write(f'{message}\n')
