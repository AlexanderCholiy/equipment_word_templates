import os
import shutil


def prepare_output_folders(output_dirs: list[str]) -> None:
    """Очистка временных папок."""
    for output_dir in output_dirs:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        else:
            for filename in os.listdir(output_dir):
                file_path = os.path.join(output_dir, filename)
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
