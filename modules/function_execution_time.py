from datetime import datetime, timedelta

from colorama import Fore, Style


def function_execution_time(func):
    """Подсчёт времени выполнения функции."""

    def wrapper(*args, **kwargs):
        start_time = datetime.now()
        try:
            result = func(*args, **kwargs)
            return result
        finally:
            end_time = datetime.now()
            execution_time = end_time - start_time
            if execution_time.total_seconds() >= 60:
                minutes = execution_time // timedelta(minutes=1)
                message = f"Функция {func.__name__} выполнялась {minutes} мин."
            elif execution_time.total_seconds() >= 1:
                seconds = round(execution_time.total_seconds(), 2)
                message = f"Функция {func.__name__} выполнялась {seconds} сек."
            else:
                microseconds: float = round(execution_time.total_seconds(), 2)
                message = (
                    f"Функция {func.__name__} выполнялась {microseconds} мкс."
                )

            print(Fore.MAGENTA + Style.DIM + message)

    return wrapper
