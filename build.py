import subprocess
import os
import shutil
import zipfile

def build_application():

    command = [
        "pyinstaller",
        "--onedir",
        "--windowed",
        "--name=GUI",
        "--icon=favicon.ico",
        "--collect-all", "tkinterdnd2",
        "-y",  # Удаляет предыдущий билд без подтверждения
        "GUI_full.py"
    ]

    print("Запуск сборки...")
    result = subprocess.run(command, capture_output=True, text=True)

    if result.returncode == 0:
        print("Сборка успешно завершена!")
        post_build_tasks()
        create_zip_archive()
    else:
        print("Ошибка при сборке:")
        print(result.stderr)


def post_build_tasks():
    dist_path = os.path.join("dist", "GUI")

    # Копирование папки makets
    source_makets = "makets"
    dest_makets = os.path.join(dist_path, "makets")
    source_resource = "resource"
    dest_resource = os.path.join(dist_path, "resource")
    dest_lst = [dest_makets, dest_resource]
    for i,source in enumerate([source_makets, source_resource]):
        if os.path.exists(source):
            print(f"Копирование {source} в {dest_lst[i]}...")
            shutil.copytree(source, dest_lst[i], dirs_exist_ok=True)
            print("Копирование завершено.")
        else:
            print(f"Папка {source} не найдена!")

    # Создание директорий /tmp_files/img и /reports/
    tmp_files_path = os.path.join(dist_path, "tmp_files", "img")
    reports_path = os.path.join(dist_path, "reports")

    for path in [tmp_files_path, reports_path]:
        if not os.path.exists(path):
            print(f"Создание директории: {path}")
            os.makedirs(path, exist_ok=True)

    print("Все дополнительные файлы и папки успешно созданы.")


def create_zip_archive():
    version = "1.1.11"
    dist_path = os.path.join("dist", "GUI")
    zip_filename = os.path.join("dist", f"GUI_{version}.zip")

    print(f"Архивация {dist_path} в {zip_filename}...")

    # Удаляем предыдущий архив, если он есть
    if os.path.exists(zip_filename):
        os.remove(zip_filename)

    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(dist_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, dist_path)
                zipf.write(file_path, arcname)

    print("Архивирование завершено.")


if __name__ == "__main__":
    build_application()
