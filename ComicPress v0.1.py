import img2pdf
import zipfile
import os
import shutil
import patoolib
import subprocess
import sys
from pathlib import Path
from natsort import natsorted 
import logging

# Настройка логирования: вывод в stdout и в файл log.txt
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler('log.txt', encoding='utf-8')
    ]
)

log = []

ALLOWED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'}

class Directory:
    @staticmethod
    def check(chapter_folder: Path):

        Page.delete_hidden(chapter_folder)

        while True:
            dirs_count = 0
            cbz_count = 0
            cbr_count = 0
            others = []

            # Сканируем содержимое директории
            for item in chapter_folder.iterdir():
                if item.is_dir():
                    dirs_count += 1
                elif item.is_file():
                    ext = item.suffix.lower()
                    if ext == '.cbz':
                        cbz_count += 1
                    elif ext == '.cbr':
                        cbr_count += 1
                    else:
                        others.append(item)
                else:
                    others.append(item)

            # Обработка посторонних элементов
            if others:
                logging.info("\nОбнаружены посторонние элементы:")
                for item in others:
                    logging.info(f"  - {item.name}")
                
                logging.info("1 - Удалить все \n2 - Прервать выполнение")
                choice = input().strip()
                if choice == '1':
                    for item in others:
                        try:
                            item.unlink()
                            logging.info(f"Удален: {item.name}")
                        except Exception as e:
                            logging.error(f"Ошибка при удалении {item.name}: {e}")
                    logging.info("Повторная проверка директории...\n")
                    continue
                else:
                    logging.info("Выполнение прервано пользователем")
                    return

            # Проверка условий после очистки
            else:
                total = dirs_count + cbz_count + cbr_count
                conditions = [
                    (dirs_count > 0 and cbz_count == 0 and cbr_count == 0),
                    (cbz_count > 0 and dirs_count == 0 and cbr_count == 0),
                    (cbr_count > 0 and dirs_count == 0 and cbz_count == 0)
                ]

                Page.delete_hidden(chapter_folder)

                if any(conditions):
                    if conditions[0]:
                        logging.info("Директория содержит только поддиректории")
                        Page.move(chapter_folder)
                    elif conditions[1]:
                        logging.info("Директория содержит только CBZ файлы")
                        CBZ.unarchiving(chapter_folder)
                    else:
                        logging.info("Директория содержит только CBR файлы")
                        CBR.unarchiving(chapter_folder)
                    return
                else:
                    logging.error("Ошибка: Директория содержит смешанные типы элементов")
                    logging.info(f"Статистика: Файлы: {dirs_count}, CBZ: {cbz_count}, CBR: {cbr_count}")
                    if total == 0:
                        logging.info("Директория пуста")
                    return



class CBZ:
    @staticmethod
    def unarchiving(chapter_folder: Path):
        # Получаем список CBZ файлов в директории
        cbz_files = list(chapter_folder.glob('*.cbz')) + list(chapter_folder.glob('*.CBZ'))
        
        for cbz_path in cbz_files:
            # Создаем имя для папки распаковки
            extract_dir = chapter_folder / cbz_path.stem
            
            try:
                # Создаем директорию для распаковки
                extract_dir.mkdir(parents=True, exist_ok=False)
                
                # Распаковываем архив
                with zipfile.ZipFile(cbz_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_dir)
                    
                # Обрабатываем вложенные директории
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        src = Path(root) / file
                        dst = extract_dir / file
                        
                        # Если файл уже в корне - пропускаем
                        if src.parent == extract_dir:
                            continue
                            
                        # Если в корне уже есть файл с таким именем — создаём уникальное имя
                        if dst.exists():
                            base = dst.stem
                            ext = dst.suffix
                            counter = 1
                            new_dst = extract_dir / f"{base}_{counter}{ext}"
                            while new_dst.exists():
                                counter += 1
                                new_dst = extract_dir / f"{base}_{counter}{ext}"
                            dst = new_dst

                        # Перемещаем файл в корневую директорию
                        try:
                            shutil.move(str(src), str(dst))
                        except Exception:
                            # В редком случае используем copy+unlink
                            shutil.copy(str(src), str(dst))
                            try:
                                src.unlink()
                            except Exception:
                                pass
                        
                # Удаляем пустые поддиректории
                for root, dirs, files in os.walk(extract_dir, topdown=False):
                    for dir in dirs:
                        dir_path = Path(root) / dir
                        try:
                            dir_path.rmdir()
                        except OSError:
                            pass
                
                # Удаляем исходный CBZ файл
                cbz_path.unlink()
                logging.info(f"Обработан: {cbz_path.name} -> {extract_dir.name}")
                
            except FileExistsError:
                logging.error(f"Ошибка: Директория {extract_dir.name} уже существует. Пропускаем.")
            except zipfile.BadZipFile:
                logging.error(f"Ошибка: Неверный формат архива {cbz_path.name}")
            except Exception as e:
                logging.exception(f"Ошибка при обработке {cbz_path.name}: {str(e)}")
                # Удаляем частично распакованные данные при ошибке
                if extract_dir.exists():
                    shutil.rmtree(extract_dir)
            Page.delete_hidden(chapter_folder)
        Page.move(chapter_folder)

    @staticmethod
    def archiving(chapter_folder: Path):
        # Проверка существования директории
        if not chapter_folder.exists() or not chapter_folder.is_dir():
            logging.error(f"Ошибка: Директория '{chapter_folder}' не существует или не является папкой")
            return

        
        # Сбор и сортировка файлов
        image_files = [
            f for f in chapter_folder.iterdir() 
            if f.is_file() and f.suffix.lower() in ALLOWED_EXTENSIONS
        ]
        
        if not image_files:
            logging.info("Нет изображений для архивации")
            return
        
        # Естественная сортировка файлов
        sorted_files = natsorted(image_files, key=lambda x: x.name)
        
        # Создание CBZ архива
        cbz_name = chapter_folder.with_name(f"{chapter_folder.name}.cbz")
        
        try:
            with zipfile.ZipFile(cbz_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for image_path in sorted_files:
                    # Добавляем файлы прямо в корень архива без путей
                    zipf.write(image_path, arcname=image_path.name)
                    
            logging.info(f"Успешно создан CBZ архив: {cbz_name}")
            
        except Exception as e:
            logging.exception(f"Ошибка при создании архива: {str(e)}")
            if cbz_name.exists():
                cbz_name.unlink()

class CBR:
    @staticmethod
    def unarchiving(chapter_folder: Path):
        if not shutil.which('rar'):
            logging.error("Ошибка: RAR не установлен или не добавлен в PATH")
            sys.exit(1)

        # Обработка CBR файлов
        for cbr_file in chapter_folder.glob('*.cbr'):
            try:
                # Создаем папку для распаковки
                extract_dir = chapter_folder / cbr_file.stem
                extract_dir.mkdir(exist_ok=True)
                
                # Распаковываем архив
                logging.info(f"Распаковываем: {cbr_file.name}")
                patoolib.extract_archive(str(cbr_file), outdir=str(extract_dir))
                
                # Обработка вложенных директорий
                for item in extract_dir.rglob('*'):
                    if item.is_file():
                        # Генерация уникального имени
                        new_path = extract_dir / item.name
                        counter = 1
                        while new_path.exists():
                            new_path = extract_dir / f"{item.stem}_{counter}{item.suffix}"
                            counter += 1
                        
                        # Перемещение файла
                        shutil.move(str(item), str(new_path))
                        
                        # Удаление пустых директорий
                        for parent in item.parents:
                            if parent != extract_dir and parent.is_dir():
                                try:
                                    parent.rmdir()
                                except OSError:
                                    break
                
                # Удаляем исходный CBR
                cbr_file.unlink()
                logging.info(f"Успешно обработан: {cbr_file.name}")
                
            except Exception as e:
                logging.exception(f"Ошибка при обработке {cbr_file.name}: {str(e)}")
                if extract_dir.exists():
                    shutil.rmtree(extract_dir)
        Page.delete_hidden(chapter_folder)
        Page.move(chapter_folder)

    @staticmethod
    def archiving(source_dir: Path):
        if not shutil.which('rar'):
            logging.error("Ошибка: RAR не установлен или не добавлен в PATH")
            sys.exit(1)

        # Проверка существования директории
        if not source_dir.exists() or not source_dir.is_dir():
            logging.error(f"Ошибка: Директория {source_dir} не существует")
            return

        # Получаем и сортируем изображения

        files = [f for f in source_dir.iterdir() if f.is_file() and f.suffix.lower() in ALLOWED_EXTENSIONS]
        
        if not files:
            logging.info("Нет изображений для архивации")
            return

        sorted_files = natsorted(files, key=lambda x: x.name)

        # Создаем временную папку
        temp_dir = source_dir / "~temp_cbr"
        temp_dir.mkdir(exist_ok=True)

        try:
            # Копируем файлы с сохранением порядка во временную папку
            for idx, file in enumerate(sorted_files, 1):
                new_name = f"{idx:04d}{file.suffix}"
                shutil.copy(file, temp_dir / new_name)

            # Создаем архив
            archive_name = source_dir.parent / f"{source_dir.name}.cbr"

            # Составляем список файлов для rar (абсолютные пути), сортируем по имени
            file_list = [str(p) for p in natsorted([p for p in temp_dir.iterdir() if p.is_file()], key=lambda x: x.name)]

            rar_cmd = ['rar', 'a', '-ep', '-idq', str(archive_name)] + file_list

            # Выполняем команду RAR
            result = subprocess.run(rar_cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                logging.error("Ошибка при создании архива:")
                logging.error(result.stderr)
                if archive_name.exists():
                    archive_name.unlink()
            else:
                logging.info(f"Архив успешно создан: {archive_name}")
                if log:
                    with open('log.txt', 'w') as f:
                        f.writelines(line + '\n' for line in log)

        finally:
            # Удаляем временную папку
            shutil.rmtree(temp_dir, ignore_errors=True)


class Page:
    @staticmethod
    def delete_hidden(root_dir: Path):
        """Рекурсивно удаляет скрытые файлы (имена, начинающиеся с '.' или '._')."""
        if not root_dir or not root_dir.exists() or not root_dir.is_dir():
            return

        for item in root_dir.rglob('*'):
            try:
                if item.is_file() and (item.name.startswith('._') or item.name.startswith('.')):
                    item.unlink()
                    logging.info(f"Удален скрытый файл: {item}")
            except Exception as e:
                logging.error(f"Ошибка при удалении {item}: {e}")

    @staticmethod
    def move(chapter_folder: Path):
        if not chapter_folder or not chapter_folder.exists() or not chapter_folder.is_dir():
            logging.error(f"Ошибка: '{chapter_folder}' не является допустимой директорией")
            return

        Page.delete_hidden(chapter_folder)

        # Получаем отсортированный список глав
        chapters = natsorted([chap for chap in chapter_folder.iterdir() if chap.is_dir()])

        all_images = []

        # Собираем изображения из глав, удаляем посторонние файлы/папки
        for chapter in chapters:
            Page.delete_hidden(chapter)

            for item in chapter.iterdir():
                if item.is_file():
                    if item.suffix.lower() not in ALLOWED_EXTENSIONS:
                        try:
                            item.unlink()
                            logging.info(f"Удаление постороннего файла: {item}")
                        except Exception as e:
                            logging.error(f"Ошибка при удалении {item}: {e}")
                elif item.is_dir():
                    try:
                        shutil.rmtree(item)
                        logging.info(f"Удаление поддиректории: {item}")
                    except Exception as e:
                        logging.error(f"Ошибка при удалении папки {item}: {e}")

            images = natsorted([img for img in chapter.iterdir() if img.is_file() and img.suffix.lower() in ALLOWED_EXTENSIONS], key=lambda x: x.name)
            all_images.extend([(chapter, img) for img in images])

        # Переименовываем и перемещаем файлы в корень папки главы
        for idx, (chapter, img_path) in enumerate(all_images, 1):
            extension = img_path.suffix
            new_name = f"{idx:04d}{extension}"
            new_path = chapter_folder / new_name

            # Генерируем уникальное имя если конфликт
            if new_path.exists():
                base = new_path.stem
                counter = 1
                candidate = chapter_folder / f"{base}_{counter}{extension}"
                while candidate.exists():
                    counter += 1
                    candidate = chapter_folder / f"{base}_{counter}{extension}"
                new_path = candidate

            try:
                img_path.replace(new_path)
                logging.info(f"Глава: {chapter.name}, Страница: {img_path.name} -> Номер: {new_path.name}")
            except Exception as e:
                logging.error(f"Ошибка при перемещении {img_path}: {e}")

        # Удаляем пустые папки глав
        for chapter in chapters:
            try:
                shutil.rmtree(chapter)
                logging.info(f"Удаление папки главы: {chapter.name}")
            except Exception as e:
                logging.error(f"Ошибка при удалении главы {chapter.name}: {e}")

        Page.choosing_action(chapter_folder)

    @staticmethod
    def choosing_action(chapter_folder: Path):
        logging.info(f"Перед конвертацией вы можете вставить обложку с именем '0' в папку {chapter_folder}")
        while True:
            logging.info('1 - PDF \n2 - CBZ \n3 - CBR')
            ans = input('Выберите формат: ').strip()
            if ans == '1':   # PDF
                PDF.create(chapter_folder)
                break
            elif ans == '2': # CBZ
                CBZ.archiving(chapter_folder)
                break
            elif ans == '3': # CBR
                CBR.archiving(chapter_folder)
                break
            else:
                logging.info('Неверный ввод, попробуйте ещё раз.')


class PDF:
    @staticmethod
    def create(chapter_folder: Path):
        # Проверяем валидность директории
        if not chapter_folder.is_dir():
            logging.error(f"Ошибка: {chapter_folder} не является директорией")
            return

        # Собираем изображения
        images = [f for f in chapter_folder.iterdir() 
                if f.is_file() and f.suffix.lower() in ALLOWED_EXTENSIONS]
        
        if not images:
            logging.info("В директории нет подходящих изображений")
            return

        # Сортируем естественным образом
        sorted_images = natsorted(images, key=lambda x: x.name)
        
        # Создаем PDF
        pdf_name = chapter_folder.parent / f"{chapter_folder.name}.pdf"
        try:
            with open(pdf_name, "wb") as f:
                f.write(img2pdf.convert([str(img) for img in sorted_images]))
            logging.info(f"PDF успешно создан: {pdf_name}")
        except Exception as e:
            logging.exception(f"Ошибка при создании PDF: {str(e)}")
            if pdf_name.exists():
                pdf_name.unlink()



if __name__ == "__main__":
    # Укажите путь к проверяемой директории
    chapter_folder = '/Users/lev/Книги/Домекано'

    
    
    logging.info('Укажите путь к проверяемой директории: ')
    chapter_folder = Path(chapter_folder)
    logging.info(f"Перед конвертацией вы можете вставить обложку с именем '0' в папку {chapter_folder}")

    #  Разкоментить необходимое (дефолт - 2)

    # 1 если директория заведомо содержит только изображения
    
    # Page.move(chapter_folder) 




    # 2 проверка что указана именно директория с главами (дефолтный вариант)
 
    if not chapter_folder.exists() or not chapter_folder.is_dir():
        logging.error("Ошибка: Указанная директория не существует или не является папкой")
    else:
        Directory.check(chapter_folder)

    Page.choosing_action(chapter_folder)
