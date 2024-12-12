import os


def prepare_folders(directories: list[str]):
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
