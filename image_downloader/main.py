import os
import sys
import threading
import queue
from pathlib import Path
from datetime import datetime
import pandas as pd
import requests
from PIL import Image
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from urllib.parse import urlparse
import hashlib
import io
import logging
import re

class ImageDownloaderApp:
    """Профессиональное приложение для скачивания и конвертации изображений"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Image Downloader Pro")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # Настройка стилей
        self.setup_styles()
        
        # Переменные
        self.excel_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=str(Path.home() / "Downloads" / "images"))
        self.url_column = tk.StringVar(value="URL")
        self.name_column = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Готов к работе")
        self.progress_var = tk.DoubleVar()
        
        # Данные Excel
        self.df = None
        self.available_columns = []
        
        # Очередь для логирования
        self.log_queue = queue.Queue()
        
        # Флаги
        self.is_running = False
        self.stop_requested = False
        
        # Создание GUI
        self.create_widgets()
        
        # Настройка логирования
        self.setup_logging()
        
        # Обработка закрытия окна
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Запуск обработчика очереди логов
        self.process_log_queue()
        
    def setup_styles(self):
        """Настройка стилей интерфейса"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Настройка цветов
        style.configure('Title.TLabel', font=('Arial', 14, 'bold'))
        style.configure('Status.TLabel', font=('Arial', 10))
        style.configure('Download.TButton', font=('Arial', 11, 'bold'))
        style.configure('Info.TLabel', font=('Arial', 9))
        
    def create_widgets(self):
        """Создание элементов интерфейса"""
        # Главный контейнер
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка сетки
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Image Downloader Pro", 
                                style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Выбор Excel файла
        ttk.Label(main_frame, text="Excel файл:").grid(
            row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_path, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(main_frame, text="Обзор...", 
                  command=self.browse_excel).grid(row=1, column=2, pady=5)
        
        # Кнопка загрузки колонок
        ttk.Button(main_frame, text="Загрузить колонки", 
                  command=self.load_columns).grid(
            row=2, column=1, sticky=tk.E, pady=5)
        
        # Выбор папки для сохранения
        ttk.Label(main_frame, text="Папка сохранения:").grid(
            row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(
            row=3, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(main_frame, text="Обзор...", 
                  command=self.browse_output).grid(row=3, column=2, pady=5)
        
        # Выбор колонок
        columns_frame = ttk.LabelFrame(main_frame, text="Настройка колонок", 
                                      padding="10")
        columns_frame.grid(row=4, column=0, columnspan=3, 
                          sticky=(tk.W, tk.E), pady=10)
        columns_frame.columnconfigure(1, weight=1)
        
        # Колонка с URL
        ttk.Label(columns_frame, text="Колонка с URL:").grid(
            row=0, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.url_column_combo = ttk.Combobox(columns_frame, 
                                            textvariable=self.url_column,
                                            state="readonly", width=30)
        self.url_column_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        self.url_column_combo.bind('<<ComboboxSelected>>', 
                                  self.on_url_column_selected)
        
        # Колонка с именами
        ttk.Label(columns_frame, text="Колонка с именами:").grid(
            row=1, column=0, sticky=tk.W, pady=5, padx=(0, 10))
        self.name_column_combo = ttk.Combobox(columns_frame, 
                                             textvariable=self.name_column,
                                             state="readonly", width=30)
        self.name_column_combo.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        self.name_column_combo.bind('<<ComboboxSelected>>', 
                                   self.on_name_column_selected)
        
        # Информация о выборе
        self.info_label = ttk.Label(columns_frame, 
                                   text="Выберите Excel файл и нажмите 'Загрузить колонки'",
                                   style='Info.TLabel', foreground='gray')
        self.info_label.grid(row=2, column=0, columnspan=2, 
                            sticky=tk.W, pady=(10, 0))
        
        # Дополнительные опции
        options_frame = ttk.LabelFrame(main_frame, text="Опции", padding="5")
        options_frame.grid(row=5, column=0, columnspan=3, 
                          sticky=(tk.W, tk.E), pady=10)
        
        self.create_subdirs_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Создавать подпапки по строкам",
                       variable=self.create_subdirs_var).grid(
            row=0, column=0, sticky=tk.W, padx=5)
        
        self.skip_existing_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Пропускать существующие файлы",
                       variable=self.skip_existing_var).grid(
            row=0, column=1, sticky=tk.W, padx=5)
        
        self.clean_names_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Очищать имена от недопустимых символов",
                       variable=self.clean_names_var).grid(
            row=0, column=2, sticky=tk.W, padx=5)
        
        # Прогресс бар
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=6, column=0, columnspan=3, 
                           sticky=(tk.W, tk.E), pady=10)
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, 
                                           variable=self.progress_var,
                                           maximum=100)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, columnspan=3, pady=10)
        
        self.start_button = ttk.Button(button_frame, text="Начать загрузку",
                                      command=self.start_download,
                                      style='Download.TButton')
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(button_frame, text="Остановить",
                                     command=self.stop_download,
                                     state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Статус
        ttk.Label(main_frame, textvariable=self.status_var,
                 style='Status.TLabel').grid(
            row=8, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        # Лог
        log_frame = ttk.LabelFrame(main_frame, text="Лог операций", padding="5")
        log_frame.grid(row=9, column=0, columnspan=3, 
                      sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, 
                                                  wrap=tk.WORD,
                                                  state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка весов для resizable
        main_frame.rowconfigure(9, weight=1)
        
    def setup_logging(self):
        """Настройка логирования"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('image_downloader.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        
    def process_log_queue(self):
        """Обработка очереди логов"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
                self.log_text.configure(state=tk.DISABLED)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_log_queue)
            
    def log(self, message, level="INFO"):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_queue.put(formatted_message)
        
        if level == "ERROR":
            self.logger.error(message)
        elif level == "WARNING":
            self.logger.warning(message)
        else:
            self.logger.info(message)
            
    def browse_excel(self):
        """Выбор Excel файла"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), 
                      ("All files", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
            self.load_columns()  # Автоматически загружаем колонки
            
    def load_columns(self):
        """Загрузка списка колонок из Excel файла"""
        if not self.excel_path.get():
            messagebox.showwarning("Предупреждение", "Сначала выберите Excel файл!")
            return
            
        try:
            # Чтение Excel файла
            self.df = pd.read_excel(self.excel_path.get())
            self.available_columns = list(self.df.columns)
            
            # Обновление комбобоксов
            self.url_column_combo['values'] = self.available_columns
            self.name_column_combo['values'] = [''] + self.available_columns
            
            # Автоматический выбор колонок
            if self.available_columns:
                # Пытаемся найти колонку с URL
                url_keywords = ['url', 'link', 'ссылка', 'адрес', 'image_url', 
                               'фото', 'изображение']
                name_keywords = ['name', 'имя', 'название', 'title', 'заголовок',
                                'filename', 'имя_файла']
                
                # Поиск колонки URL
                url_col = None
                for keyword in url_keywords:
                    for col in self.available_columns:
                        if keyword.lower() in col.lower():
                            url_col = col
                            break
                    if url_col:
                        break
                        
                # Поиск колонки с именами
                name_col = None
                for keyword in name_keywords:
                    for col in self.available_columns:
                        if keyword.lower() in col.lower() and col != url_col:
                            name_col = col
                            break
                    if name_col:
                        break
                
                # Установка найденных колонок
                if url_col:
                    self.url_column.set(url_col)
                elif self.available_columns:
                    self.url_column.set(self.available_columns[0])
                    
                if name_col:
                    self.name_column.set(name_col)
                    
                # Обновление информации
                info_text = f"Найдено колонок: {len(self.available_columns)}"
                if url_col:
                    info_text += f"\nКолонка URL: {url_col}"
                if name_col:
                    info_text += f"\nКолонка имен: {name_col}"
                if not name_col:
                    info_text += "\nКолонка имен не найдена - будут использованы автоматические имена"
                    
                self.info_label.config(text=info_text, foreground='green')
                self.log(f"Загружены колонки: {', '.join(self.available_columns)}")
                
        except Exception as e:
            self.log(f"Ошибка чтения Excel: {str(e)}", "ERROR")
            messagebox.showerror("Ошибка", f"Не удалось прочитать Excel файл:\n{str(e)}")
            
    def on_url_column_selected(self, event):
        """Обработка выбора колонки URL"""
        selected = self.url_column.get()
        self.log(f"Выбрана колонка URL: {selected}")
        
    def on_name_column_selected(self, event):
        """Обработка выбора колонки имен"""
        selected = self.name_column.get()
        if selected:
            self.log(f"Выбрана колонка имен: {selected}")
        else:
            self.log("Колонка имен не выбрана - будут использованы автоматические имена")
            
    def browse_output(self):
        """Выбор папки для сохранения"""
        directory = filedialog.askdirectory(
            title="Выберите папку для сохранения",
            initialdir=self.output_dir.get()
        )
        if directory:
            self.output_dir.set(directory)
            
    def validate_inputs(self):
        """Проверка входных данных"""
        if not self.excel_path.get():
            messagebox.showerror("Ошибка", "Выберите Excel файл!")
            return False
            
        if not os.path.exists(self.excel_path.get()):
            messagebox.showerror("Ошибка", "Excel файл не существует!")
            return False
            
        if not self.url_column.get().strip():
            messagebox.showerror("Ошибка", "Выберите колонку с URL!")
            return False
            
        if self.df is None or self.url_column.get() not in self.df.columns:
            messagebox.showerror("Ошибка", "Колонка URL не найдена в файле!")
            return False
            
        if self.name_column.get() and self.name_column.get() not in self.df.columns:
            messagebox.showerror("Ошибка", "Колонка имен не найдена в файле!")
            return False
            
        return True
        
    def clean_filename(self, filename):
        """Очистка имени файла от недопустимых символов"""
        if not self.clean_names_var.get():
            return filename
            
        # Удаление недопустимых символов
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        # Удаление пробелов в начале и конце
        filename = filename.strip()
        # Замена множественных пробелов на один
        filename = re.sub(r'\s+', ' ', filename)
        # Удаление точек в конце (Windows)
        filename = filename.rstrip('.')
        
        return filename if filename else 'unnamed'
        
    def get_image_data(self):
        """Получение данных изображений из Excel"""
        try:
            url_column = self.url_column.get().strip()
            name_column = self.name_column.get().strip() if self.name_column.get() else None
            
            # Получение URL
            urls = self.df[url_column].dropna().astype(str).tolist()
            
            # Фильтрация валидных URL
            valid_indices = []
            valid_urls = []
            
            for i, url in enumerate(urls):
                url = url.strip()
                if url.startswith(('http://', 'https://')):
                    valid_indices.append(i)
                    valid_urls.append(url)
                    
            if not valid_urls:
                raise ValueError("В колонке не найдено валидных URL!")
                
            # Получение имен
            names = []
            if name_column:
                raw_names = self.df[name_column].tolist()
                for i in valid_indices:
                    if i < len(raw_names) and pd.notna(raw_names[i]):
                        name = self.clean_filename(str(raw_names[i]))
                        if name and name != 'nan':
                            names.append(name)
                        else:
                            names.append(None)
                    else:
                        names.append(None)
            else:
                names = [None] * len(valid_urls)
                
            self.log(f"Найдено {len(valid_urls)} валидных URL")
            if name_column:
                named_count = sum(1 for n in names if n)
                self.log(f"Из них с именами: {named_count}")
                
            return valid_urls, names
            
        except Exception as e:
            self.log(f"Ошибка чтения данных: {str(e)}", "ERROR")
            messagebox.showerror("Ошибка", f"Не удалось прочитать данные:\n{str(e)}")
            return [], []
            
    def generate_filename(self, url, custom_name, index):
        """Генерация имени файла"""
        if custom_name:
            # Используем пользовательское имя
            filename = custom_name
        else:
            # Автоматическая генерация имени
            path = urlparse(url).path
            original_name = os.path.basename(path)
            
            if not original_name or '.' not in original_name:
                hash_object = hashlib.md5(url.encode())
                original_name = f"image_{hash_object.hexdigest()[:10]}"
                
            name_without_ext = os.path.splitext(original_name)[0]
            if not name_without_ext:
                name_without_ext = f"image_{index}"
                
            filename = name_without_ext
            
        # Очистка имени
        filename = self.clean_filename(filename)
        
        # Добавление расширения
        return f"{filename}.png"
        
    def download_and_convert(self, url, output_path):
        """Скачивание и конвертация изображения"""
        try:
            # Настройка заголовков для имитации браузера
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            # Скачивание
            response = requests.get(url, headers=headers, timeout=30, 
                                   allow_redirects=True)
            response.raise_for_status()
            
            # Проверка типа контента
            content_type = response.headers.get('content-type', '')
            if not content_type.startswith('image/'):
                raise ValueError(f"URL не содержит изображение: {content_type}")
                
            # Открытие изображения
            image = Image.open(io.BytesIO(response.content))
            
            # Конвертация в RGB (если нужно)
            if image.mode in ('RGBA', 'LA', 'P'):
                if image.mode == 'P':
                    image = image.convert('RGBA')
                # Создаем белый фон для прозрачных изображений
                background = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'RGBA':
                    background.paste(image, mask=image.split()[3])
                else:
                    background.paste(image)
                image = background
            elif image.mode != 'RGB':
                image = image.convert('RGB')
                
            # Сохранение как PNG
            image.save(output_path, 'PNG', optimize=True)
            
            return True
            
        except requests.exceptions.RequestException as e:
            self.log(f"Ошибка скачивания {url}: {str(e)}", "ERROR")
            return False
        except Exception as e:
            self.log(f"Ошибка обработки {url}: {str(e)}", "ERROR")
            return False
            
    def start_download(self):
        """Запуск процесса загрузки"""
        if self.is_running:
            return
            
        if not self.validate_inputs():
            return
            
        # Получение данных
        urls, names = self.get_image_data()
        if not urls:
            return
            
        # Создание папки
        output_dir = Path(self.output_dir.get())
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Сброс флагов
        self.is_running = True
        self.stop_requested = False
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_var.set(0)
        
        # Запуск в отдельном потоке
        thread = threading.Thread(target=self.download_worker, 
                                 args=(urls, names, output_dir))
        thread.daemon = True
        thread.start()
        
    def download_worker(self, urls, names, output_dir):
        """Рабочий поток для загрузки"""
        total_urls = len(urls)
        successful = 0
        failed = 0
        skipped = 0
        renamed = 0
        
        self.log(f"Начало загрузки {total_urls} изображений")
        self.status_var.set(f"Загрузка: 0/{total_urls}")
        
        for i, (url, name) in enumerate(zip(urls, names), 1):
            if self.stop_requested:
                self.log("Загрузка остановлена пользователем", "WARNING")
                break
                
            try:
                # Генерация пути для сохранения
                filename = self.generate_filename(url, name, i)
                
                # Создание подпапки если нужно
                if self.create_subdirs_var.get():
                    subdir = output_dir / f"row_{i:04d}"
                    subdir.mkdir(exist_ok=True)
                else:
                    subdir = output_dir
                    
                output_path = subdir / filename
                
                # Проверка на дубликаты имен
                if output_path.exists() and not self.skip_existing_var.get():
                    # Если файл существует и не нужно пропускать, добавляем номер
                    base_name = filename.rsplit('.', 1)[0]
                    counter = 1
                    while output_path.exists():
                        filename = f"{base_name}_{counter}.png"
                        output_path = subdir / filename
                        counter += 1
                    renamed += 1
                    
                # Проверка существующего файла
                if self.skip_existing_var.get() and output_path.exists():
                    self.log(f"Пропущен (существует): {filename}")
                    skipped += 1
                else:
                    # Скачивание и конвертация
                    name_info = f" (имя: {name})" if name else ""
                    self.log(f"Загрузка [{i}/{total_urls}]: {url}{name_info}")
                    
                    if self.download_and_convert(url, output_path):
                        self.log(f"Успешно сохранен: {output_path}")
                        successful += 1
                    else:
                        failed += 1
                        
            except Exception as e:
                self.log(f"Ошибка обработки {url}: {str(e)}", "ERROR")
                failed += 1
                
            # Обновление прогресса
            progress = (i / total_urls) * 100
            self.progress_var.set(progress)
            self.status_var.set(f"Загрузка: {i}/{total_urls} "
                              f"(Успешно: {successful}, Ошибок: {failed}, "
                              f"Пропущено: {skipped})")
            
        # Завершение
        self.is_running = False
        self.root.after(0, self.download_complete, successful, failed, 
                       skipped, renamed)
        
    def download_complete(self, successful, failed, skipped, renamed):
        """Обработка завершения загрузки"""
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.progress_var.set(100)
        
        summary = f"Загрузка завершена! Успешно: {successful}, " \
                 f"Ошибок: {failed}, Пропущено: {skipped}"
        if renamed:
            summary += f", Переименовано: {renamed}"
            
        self.status_var.set(summary)
        self.log(summary, "INFO" if failed == 0 else "WARNING")
        
        messagebox.showinfo("Готово", summary)
        
    def stop_download(self):
        """Остановка загрузки"""
        self.stop_requested = True
        self.log("Запрос на остановку загрузки...", "WARNING")
        
    def on_closing(self):
        """Обработка закрытия окна"""
        if self.is_running:
            if messagebox.askokcancel("Выход", 
                                     "Загрузка еще выполняется. Выйти?"):
                self.stop_requested = True
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    """Главная функция"""
    root = tk.Tk()
    app = ImageDownloaderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()