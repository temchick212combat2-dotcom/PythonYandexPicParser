import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from parser import YandexImagesLinkParser  # Импортируем новый парсер
import threading
import os
import pandas as pd
from datetime import datetime

class YandexImagesParserGUI:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("📸 Парсер Яндекс Картинок - Получение ссылок")
        self.window.geometry("1050x800")
        self.window.resizable(True, True)
        
        self.colors = {
            'bg': '#2b2b2b',
            'fg': '#ffffff',
            'button': '#3a7ebf',
            'button_hover': '#4a8edf',
            'button_success': '#2ea043',
            'button_danger': '#da3633',
            'entry': '#3a3a3a',
            'frame': '#1a1a1a',
            'warning': '#ffaa00',
            'success': '#00ff88',
            'info': '#58a6ff'
        }
        
        self.window.configure(bg=self.colors['bg'])
        
        # Переменные
        self.excel_file_path = tk.StringVar()
        self.column_name = tk.StringVar(value="")
        self.limit_items = tk.StringVar(value="10")
        self.start_from = tk.StringVar(value="0")
        self.delay_between = tk.StringVar(value="5")
        self.output_file = tk.StringVar(value="image_links.csv")
        self.is_running = False
        self.parser = None
        
        self.setup_ui()
        
    def create_button(self, parent, text, command=None, **kwargs):
        """Создание стилизованной кнопки"""
        btn = tk.Button(
            parent,
            text=text,
            command=command,
            bg=self.colors['button'],
            fg='white',
            font=("Arial", 10, "bold"),
            relief=tk.FLAT,
            padx=10,
            pady=5,
            cursor="hand2",
            **kwargs
        )
        
        def on_enter(e):
            btn.configure(bg=self.colors['button_hover'])
        
        def on_leave(e):
            btn.configure(bg=self.colors['button'])
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn
    
    def setup_ui(self):
        main_frame = tk.Frame(self.window, bg=self.colors['bg'])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Заголовок
        title_label = tk.Label(
            main_frame,
            text="📸 Парсер Яндекс Картинок - Получение ссылок",
            font=("Arial", 26, "bold"),
            bg=self.colors['bg'],
            fg=self.colors['fg']
        )
        title_label.pack(pady=(0, 5))
        
        subtitle_label = tk.Label(
            main_frame,
            text="Получает ссылки на изображения по поисковым запросам",
            font=("Arial", 12),
            bg=self.colors['bg'],
            fg='#888888'
        )
        subtitle_label.pack(pady=(0, 15))
        
        # Предупреждение
        warning_label = tk.Label(
            main_frame,
            text="⚠️ Для избежания каптчи установите задержку 5-10 секунд",
            bg=self.colors['bg'],
            fg=self.colors['warning'],
            font=("Arial", 11, "bold")
        )
        warning_label.pack(pady=(0, 15))
        
        # Фрейм для выбора файла
        file_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        file_frame.pack(fill="x", pady=(0, 10))
        
        file_label = tk.Label(
            file_frame,
            text="📂 Excel файл:",
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 12)
        )
        file_label.pack(side="left", padx=(0, 10))
        
        file_entry = tk.Entry(
            file_frame,
            textvariable=self.excel_file_path,
            bg=self.colors['entry'],
            fg=self.colors['fg'],
            font=("Arial", 10),
            relief=tk.FLAT,
            width=50
        )
        file_entry.pack(side="left", padx=(0, 10), fill="x", expand=True)
        
        browse_btn = self.create_button(
            file_frame,
            text="Обзор",
            command=self.browse_excel,
            width=10
        )
        browse_btn.pack(side="left")
        
        # Настройки
        settings_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        settings_frame.pack(fill="x", pady=(0, 10))
        
        # Столбец
        column_label = tk.Label(
            settings_frame,
            text="Столбец:",
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 10)
        )
        column_label.pack(side="left", padx=(0, 5))
        
        column_entry = tk.Entry(
            settings_frame,
            textvariable=self.column_name,
            bg=self.colors['entry'],
            fg=self.colors['fg'],
            font=("Arial", 10),
            relief=tk.FLAT,
            width=15
        )
        column_entry.pack(side="left", padx=(0, 15))
        
        # Лимит
        limit_label = tk.Label(
            settings_frame,
            text="Обработать:",
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 10)
        )
        limit_label.pack(side="left", padx=(0, 5))
        
        limit_entry = tk.Entry(
            settings_frame,
            textvariable=self.limit_items,
            bg=self.colors['entry'],
            fg=self.colors['fg'],
            font=("Arial", 10),
            relief=tk.FLAT,
            width=8
        )
        limit_entry.pack(side="left", padx=(0, 15))
        
        # Задержка
        delay_label = tk.Label(
            settings_frame,
            text="Задержка (сек):",
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 10)
        )
        delay_label.pack(side="left", padx=(0, 5))
        
        delay_entry = tk.Entry(
            settings_frame,
            textvariable=self.delay_between,
            bg=self.colors['entry'],
            fg=self.colors['fg'],
            font=("Arial", 10),
            relief=tk.FLAT,
            width=8
        )
        delay_entry.pack(side="left", padx=(0, 15))
        
        # Смещение
        start_label = tk.Label(
            settings_frame,
            text="Начать с:",
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 10)
        )
        start_label.pack(side="left", padx=(0, 5))
        
        start_entry = tk.Entry(
            settings_frame,
            textvariable=self.start_from,
            bg=self.colors['entry'],
            fg=self.colors['fg'],
            font=("Arial", 10),
            relief=tk.FLAT,
            width=8
        )
        start_entry.pack(side="left", padx=(0, 15))
        
        # Имя выходного файла
        output_label = tk.Label(
            settings_frame,
            text="Сохранить в:",
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 10)
        )
        output_label.pack(side="left", padx=(0, 5))
        
        output_entry = tk.Entry(
            settings_frame,
            textvariable=self.output_file,
            bg=self.colors['entry'],
            fg=self.colors['fg'],
            font=("Arial", 10),
            relief=tk.FLAT,
            width=12
        )
        output_entry.pack(side="left")
        
        # Информация
        info_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        info_frame.pack(fill="x", pady=(0, 10))
        
        info_label = tk.Label(
            info_frame,
            text="📁 Результат: CSV файл | 💡 0 = все позиции | ⚠️ Рекомендуемая задержка: 5-10 сек",
            bg=self.colors['bg'],
            fg='#888888',
            font=("Arial", 10)
        )
        info_label.pack(side="left")
        
        # Кнопки
        button_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        button_frame.pack(fill="x", pady=(0, 10))
        
        self.start_btn = self.create_button(
            button_frame,
            text="▶ Начать парсинг",
            command=self.start_parsing,
            width=20,
            height=2
        )
        self.start_btn.configure(font=("Arial", 12, "bold"))
        self.start_btn.pack(side="left", padx=(0, 10), expand=True, fill="x")
        
        self.stop_btn = self.create_button(
            button_frame,
            text="⏹ Остановить",
            command=self.stop_parsing,
            width=20,
            height=2,
            state="disabled"
        )
        self.stop_btn.configure(font=("Arial", 12, "bold"))
        self.stop_btn.pack(side="left", expand=True, fill="x")
        
        self.view_btn = self.create_button(
            button_frame,
            text="📊 Просмотр CSV",
            command=self.view_results,
            width=15,
            height=2
        )
        self.view_btn.configure(font=("Arial", 12, "bold"))
        self.view_btn.pack(side="left", padx=(10, 0), expand=True, fill="x")
        
        # Прогресс
        progress_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        progress_frame.pack(fill="x", pady=(0, 10))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.pack(fill="x")
        
        # Статус
        self.status_label = tk.Label(
            main_frame,
            text="✅ Готов к работе",
            bg=self.colors['bg'],
            fg=self.colors['success'],
            font=("Arial", 11)
        )
        self.status_label.pack(pady=(5, 0))
        
        # Лог
        log_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        log_frame.pack(fill="both", expand=True, pady=(10, 0))
        
        log_label = tk.Label(
            log_frame,
            text="📋 Лог работы:",
            bg=self.colors['bg'],
            fg=self.colors['fg'],
            font=("Arial", 12, "bold")
        )
        log_label.pack(anchor="w", pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=20,
            bg='#1a1a1a',
            fg='#00ff00',
            font=("Consolas", 9),
            wrap=tk.WORD
        )
        self.log_text.pack(fill="both", expand=True)
        
        self.log_text.tag_config("INFO", foreground="#00ff00")
        self.log_text.tag_config("WARNING", foreground="#ffaa00")
        self.log_text.tag_config("ERROR", foreground="#ff4444")
        self.log_text.tag_config("SUCCESS", foreground="#00ff88")
        
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file_path.set(filename)
            self.log_message(f"✅ Выбран файл: {os.path.basename(filename)}", "SUCCESS")
            self.preview_excel(filename)
    
    def preview_excel(self, filename):
        try:
            df = pd.read_excel(filename, nrows=5)
            self.log_message("=" * 50, "SUCCESS")
            self.log_message(f"📊 Столбцы: {', '.join(df.columns)}", "INFO")
            self.log_message(f"📏 Показано первых {len(df)} строк", "INFO")
            self.log_message("=" * 50, "SUCCESS")
        except Exception as e:
            self.log_message(f"❌ Ошибка: {str(e)}", "ERROR")
    
    def log_message(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n", level)
        self.log_text.see("end")
        self.window.update()
    
    def update_progress(self, current, total):
        progress = (current / total * 100) if total > 0 else 0
        self.progress_var.set(progress)
        self.window.update()
    
    def start_parsing(self):
        if self.is_running:
            return
        
        if not self.excel_file_path.get():
            messagebox.showerror("Ошибка", "Выберите Excel файл")
            return
        
        self.is_running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.status_label.configure(text="🔄 Выполняется парсинг...", fg=self.colors['warning'])
        
        self.log_message("=" * 50, "SUCCESS")
        self.log_message("🚀 Начинаем парсинг (получение ссылок)...", "SUCCESS")
        self.log_message(f"📂 Файл: {os.path.basename(self.excel_file_path.get())}", "INFO")
        self.log_message(f"📁 Результат: {self.output_file.get()}", "INFO")
        
        limit = int(self.limit_items.get()) if self.limit_items.get().strip() else 0
        start = int(self.start_from.get()) if self.start_from.get().strip() else 0
        delay = int(self.delay_between.get()) if self.delay_between.get().strip() else 5
        
        if limit > 0:
            self.log_message(f"📊 Обработать: {limit} позиций", "INFO")
        if start > 0:
            self.log_message(f"📊 Начать с: {start}", "INFO")
        self.log_message(f"⏱ Задержка между запросами: {delay} сек", "INFO")
        self.log_message("⚠️ Для избежания каптчи соблюдайте задержки!", "WARNING")
        
        threading.Thread(target=self.run_parsing, daemon=True).start()
    
    def run_parsing(self):
        try:
            # Создаем парсер для получения ссылок
            self.parser = YandexImagesLinkParser(output_file=self.output_file.get())
            
            self.log_message("🌐 Настройка браузера...", "INFO")
            if not self.parser.setup_driver():
                self.log_message("❌ Не удалось настроить браузер", "ERROR")
                self.finish_parsing()
                return
            
            column = self.column_name.get().strip() if self.column_name.get().strip() else None
            limit = int(self.limit_items.get()) if self.limit_items.get().strip() else 0
            start = int(self.start_from.get()) if self.start_from.get().strip() else 0
            
            self.log_message("📖 Чтение Excel файла...", "INFO")
            items = self.parser.parse_excel(
                self.excel_file_path.get(),
                column_name=column,
                limit=limit,
                start_from=start
            )
            
            if not items:
                self.log_message("❌ Нет данных для обработки", "ERROR")
                self.finish_parsing()
                return
            
            self.log_message(f"✅ Найдено {len(items)} позиций", "SUCCESS")
            self.update_progress(0, len(items))
            
            # Обрабатываем список
            results = self.parser.process_list(items)
            
            self.log_message("=" * 50, "SUCCESS")
            self.log_message("📊 Парсинг завершен!", "SUCCESS")
            self.log_message(f"   ✅ Успешно получено ссылок: {results['success']}", "SUCCESS")
            self.log_message(f"   ❌ Не найдено: {results['failed']}", "ERROR" if results['failed'] > 0 else "INFO")
            self.log_message(f"   📁 Результат сохранен в: {self.output_file.get()}", "INFO")
            
            # Показываем первые несколько результатов
            if results['results']:
                self.log_message("=" * 50, "SUCCESS")
                self.log_message("📋 Первые 5 полученных ссылок:", "INFO")
                for i, result in enumerate(results['results'][:5], 1):
                    self.log_message(f"   {i}. {result['query'][:40]}... -> {result['image_url'][:60]}...", "INFO")
            
        except Exception as e:
            self.log_message(f"❌ Ошибка: {str(e)}", "ERROR")
            import traceback
            self.log_message(traceback.format_exc(), "ERROR")
        finally:
            self.finish_parsing()
    
    def finish_parsing(self):
        if self.parser:
            try:
                self.parser.close()
            except:
                pass
        
        self.is_running = False
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.status_label.configure(text="✅ Готов к работе", fg=self.colors['success'])
        self.progress_var.set(100)
        self.window.after(1000, lambda: self.progress_var.set(0))
    
    def stop_parsing(self):
        if self.is_running:
            self.log_message("⚠️ Остановка парсинга...", "WARNING")
            self.is_running = False
    
    def view_results(self):
        """Просмотр результатов в CSV"""
        csv_file = self.output_file.get()
        if not os.path.exists(csv_file):
            messagebox.showinfo("Информация", f"Файл {csv_file} еще не создан.\nЗапустите парсинг для получения результатов.")
            return
        
        try:
            df = pd.read_csv(csv_file, encoding='utf-8')
            
            if df.empty:
                messagebox.showinfo("Информация", "Файл пуст. Нет сохраненных результатов.")
                return
            
            # Создаем окно для просмотра
            view_window = tk.Toplevel(self.window)
            view_window.title(f"📊 Результаты парсинга - {csv_file}")
            view_window.geometry("900x600")
            view_window.configure(bg=self.colors['bg'])
            
            # Информация
            info_label = tk.Label(
                view_window,
                text=f"Всего ссылок: {len(df)} | Файл: {csv_file}",
                bg=self.colors['bg'],
                fg=self.colors['fg'],
                font=("Arial", 12)
            )
            info_label.pack(pady=10)
            
            # Таблица
            tree_frame = tk.Frame(view_window, bg=self.colors['bg'])
            tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            tree = ttk.Treeview(
                tree_frame,
                columns=('№', 'Запрос', 'Ссылка на изображение', 'Время'),
                show='headings',
                height=25
            )
            
            tree.heading('№', text='№')
            tree.heading('Запрос', text='Запрос')
            tree.heading('Ссылка на изображение', text='Ссылка на изображение')
            tree.heading('Время', text='Время')
            
            tree.column('№', width=50, anchor='center')
            tree.column('Запрос', width=200, anchor='w')
            tree.column('Ссылка на изображение', width=500, anchor='w')
            tree.column('Время', width=150, anchor='center')
            
            scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            # Заполняем данными
            for idx, row in df.iterrows():
                tree.insert('', 'end', values=(
                    idx + 1,
                    row['query'][:50] + ('...' if len(str(row['query'])) > 50 else ''),
                    row['image_url'][:80] + ('...' if len(str(row['image_url'])) > 80 else ''),
                    row['timestamp']
                ))
            
            tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Кнопка закрытия
            close_btn = self.create_button(
                view_window,
                text="Закрыть",
                command=view_window.destroy,
                width=10
            )
            close_btn.pack(pady=10)
            
            # Двойной клик для копирования ссылки
            def copy_url(event):
                selected = tree.selection()
                if selected:
                    values = tree.item(selected[0], 'values')
                    if values and len(values) > 2:
                        self.window.clipboard_clear()
                        self.window.clipboard_append(values[2])
                        messagebox.showinfo("Успех", "Ссылка скопирована в буфер обмена!")
            
            tree.bind('<Double-Button-1>', copy_url)
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл: {str(e)}")
    
    def on_closing(self):
        if self.is_running:
            if not messagebox.askyesno("Подтверждение", "Парсинг выполняется. Вы уверены, что хотите закрыть программу?"):
                return
            self.is_running = False
        if self.parser:
            try:
                self.parser.close()
            except:
                pass
        self.window.destroy()
    
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = YandexImagesParserGUI()
    app.run()