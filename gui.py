import os
import re
import sv_ttk
import shutil
import datetime
import threading
import webbrowser
import pandas as pd
import tkinter as tk
from search_phone_sjob import VacancyParser
from search_ads_sjob import SearchSuperJob
from async_runner import AsyncParserRunner
from tkinter import ttk, messagebox, filedialog, Toplevel, Text

class SJobParser(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.parent.title("SJobParser")
        self.parent.geometry("670x700")

        try:
            self.parent.iconbitmap("static/icon.ico")
        except Exception as e:
            print(f"Cannot load icon: {e}")

        self.interface_style()
        self.pack(fill=tk.BOTH, expand=True)

        self.create_widgets()
        self.toggle_parser_mode()

        self.is_parsing = False
        self.phone_excel_path = None  # Путь к Excel файлу для парсера телефонов

        self.source_file_path = "superjob_parse_results/superjob_vacancies.xlsx"
        self.output_excel = "superjob_parse_results/superjob_vacancies_output.xlsx"
        self.clear_bugs = "sjob_phones_playwright/debug"
 
    def interface_style(self):
        sv_ttk.set_theme("light")
           
    def create_widgets(self):
        """Создание всех виджетов интерфейса"""
        self.top_level_menu()
        self.create_parser_controls()
        self.create_status_bar()
        
    def top_level_menu(self):
        """Верхнее меню"""
        menubar = tk.Menu(self.parent)
        self.parent.config(menu=menubar)

        parse_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Парсинг", menu=parse_menu)
        parse_menu.add_command(label="Открыть Excel файл...", accelerator="Ctrl+O", command=self.btn_open)
        self.parent.bind("<Control-o>", lambda _: self.btn_open())  # Горячие клавиши
        self.parent.bind("<Control-s>", lambda _: self.stop_parsing())
        self.parent.bind("<Control-l>", lambda _: self.clear_log())
        self.parent.bind("<Control-q>", lambda _: self.btn_exit())
        self.parent.bind("<Control-g>", lambda _: self.generate_url())
        self.parent.bind("<Control-r>", lambda _: self.run_parsing())
        self.parent.bind("<F1>", lambda _: self.open_link())
        parse_menu.add_separator()
        parse_menu.add_command(label="Выход", command=self.btn_exit)

        export_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Экспорт", menu=export_menu)
        export_menu.add_command(label="Экспорт объявлений...", command=self.copy_ads_file_to_path)
        export_menu.add_command(label="Экспорт готового файла...", command=self.copy_ready_file_to_path)

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Справка", menu=help_menu)
        help_menu.add_command(label="Руководство пользователя", command=self.open_link)
        help_menu.add_command(label="Горячие клавиши", command=self.hotkeys_info)
        help_menu.add_separator()
        help_menu.add_command(label="О программе", command=self.btn_about)

        # Сохраняем ссылки на меню для изменения цветов
        self.menubar = menubar
        self.parse_menu = parse_menu
        self.export_menu = export_menu
        self.help_menu = help_menu
        
    def create_parser_controls(self):
        """Создание элементов управления для парсера"""
        # Основной фрейм с grid для точного контроля
        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Конфигурация grid - основной контейнер
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Счетчик строк для grid
        row = 0
        
        # 1. Фрейм для выбора режима парсинга
        mode_frame = ttk.LabelFrame(main_frame, text="Поиск", padding=10)
        mode_frame.grid(row=row, column=0, sticky=tk.EW, padx=10, pady=(0, 5))
        mode_frame.config(height=70)
        
        self.parser_mode_key = tk.StringVar(value="keyword")
        
        ttk.Radiobutton(mode_frame, text="Поиск объявлений", 
                       variable=self.parser_mode_key, 
                       value="keyword",
                       command=self.toggle_parser_mode).grid(row=0, column=0, sticky=tk.W, padx=15, pady=0)
        
        ttk.Radiobutton(mode_frame, text="Поиск объявлений по URL", 
                       variable=self.parser_mode_key, 
                       value="url",
                       command=self.toggle_parser_mode).grid(row=0, column=1, sticky=tk.W, padx=15, pady=0)
        
        ttk.Radiobutton(mode_frame, text="Поиск телефонов",
                       variable=self.parser_mode_key,
                       value="phone",
                       command=self.toggle_parser_mode).grid(row=0, column=2, sticky=tk.W, padx=15, pady=0)

        row += 1
        
        # 2. Фрейм для темы парсера
        theme_frame = ttk.LabelFrame(main_frame, text="Тема парсера", padding=10)
        theme_frame.grid(row=row, column=0, sticky=tk.EW, padx=10, pady=(0, 5))
        theme_frame.config(height=70)
        
        self.parser_mode_t = tk.StringVar(value="tlight")
        
        ttk.Radiobutton(theme_frame, text="Светлая тема",
                       variable=self.parser_mode_t,
                       value="tlight",
                       command=self.theme_parser_mode).grid(row=0, column=0, sticky=tk.W, padx=15, pady=0)
        
        ttk.Radiobutton(theme_frame, text="Темная тема",
                       variable=self.parser_mode_t,
                       value="tdark",
                       command=self.theme_parser_mode).grid(row=0, column=1, sticky=tk.W, padx=15, pady=0)
        
        row += 1
        
        # 3. Фрейм для параметров парсинга
        self.params_frame = ttk.LabelFrame(main_frame, text="Параметры парсинга", padding=8)
        self.params_frame.grid(row=row, column=0, sticky=tk.EW, padx=10, pady=(0, 5))
        self.params_frame.config(height=90)
        
        self.create_keyword_params()
        self.create_url_params()
        self.create_phone_params()

        row += 1
        
        # 4. Дополнительные параметры
        common_frame = ttk.LabelFrame(main_frame, text="Дополнительные параметры", padding=10)
        common_frame.grid(row=row, column=0, sticky=tk.EW, padx=10, pady=(0, 5))
        common_frame.config(height=90)
        
        # Содержимое common_frame
        ttk.Label(common_frame, text="Количество фирм:").grid(row=0, column=0, sticky=tk.W, pady=0)
        self.firm_count_var = tk.IntVar(value=50)
        self.firm_count_spinbox = ttk.Spinbox(common_frame, from_=1, to=10000, 
                                              textvariable=self.firm_count_var, width=15)
        self.firm_count_spinbox.grid(row=0, column=1, padx=5, pady=0, sticky=tk.W)
        
        self.text_url_btn = ttk.Label(common_frame, text="Парсинг по URL:", width=15)
        self.text_url_btn.grid(row=1, column=0, sticky=tk.W, pady=0)
        
        self.generate_url_btn = ttk.Button(common_frame, text="Сгенерировать URL", 
                                          command=self.generate_url, width=22)
        self.generate_url_btn.grid(row=1, column=1, padx=5, pady=0, sticky=tk.W)
        
        row += 1
        
        # 5. Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=row, column=0, sticky=tk.W, padx=20, pady=4)
        button_frame.config(height=40)
        
        ttk.Button(button_frame, text="Начать парсинг", 
                    command=self.run_parsing, width=20).pack(side=tk.LEFT, padx=5)

        ttk.Button(button_frame, text="Остановить парсинг",
                    command=self.stop_parsing, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Очистить лог",
                    command=self.clear_log, width=20).pack(side=tk.LEFT, padx=5)
        
        row += 1
        
        # Лог выполнения
        log_frame = ttk.LabelFrame(main_frame, text="Лог выполнения", padding=10)
        log_frame.grid(row=row, column=0, sticky=tk.NSEW, padx=10, pady=0)
        
        # Настраиваем вес строки для растягивания лога
        main_frame.grid_rowconfigure(row, weight=1)
        
        # Создаем текстовое поле для логов
        self.log_text = tk.Text(log_frame, height=20, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Добавляем раскраску вывода текста в "Лог выполнения"
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("ERROR", foreground="red")
        self.log_text.tag_config("WARNING", foreground="#cf7c00")
        self.log_text.tag_config("SUCCESS", foreground="#00a800")
        
        # Добавляем скроллбар
        scrollbar = ttk.Scrollbar(self.log_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)
    
    def theme_parser_mode(self):
        """Переключение между темой парсера"""
        current_geometry = self.parent.geometry()  # Сохраняем текущие размеры окна
        
        if self.parser_mode_t.get() == "tlight":
            sv_ttk.set_theme("light")
            self.log_text.tag_config("INFO", foreground="black")
            self.log_text.tag_config("WARNING", foreground="#cf7c00")
            self.log_text.tag_config("SUCCESS", foreground="#00a800")
        else:
            sv_ttk.set_theme("dark")
            self.log_text.tag_config("INFO", foreground="white")
            self.log_text.tag_config("WARNING", foreground="#ffc766")
            self.log_text.tag_config("SUCCESS", foreground="#00e600")
            
        # Принудительно обновляем интерфейс
        self.parent.update_idletasks()
        
        # Восстанавливаем размеры окна
        self.parent.geometry(current_geometry)

    def generate_url(self):
        """Генерация URL на основе ключевого слова и города"""
        keyword = self.keyword_var_keyword.get().strip()
        city = self.city_var_keyword.get().strip()
        city = '%20'.join(city.split())
        if not keyword or not city:
            messagebox.showwarning("Предупреждение", "Введите ключевое слово и город!")
            return

        generated_url = f"https://www.superjob.ru/vacancy/search/?keywords={city}%20{keyword}"

        self.url_var.set(generated_url)

        # Предлагаем переключиться на режим по URL
        if messagebox.askyesno("URL сгенерирован",
                              f"URL успешно сгенерирован:\n{generated_url}\n\n"
                              f"Хотите переключиться на парсер по URL?"):
            self.parser_mode_key.set("url")
            self.toggle_parser_mode()
        self.status_var.set("URL сгенерирован")
            
    def create_keyword_params(self):
        """Создание элементов для парсера по ключу"""
        self.keyword_frame = ttk.Frame(self.params_frame)
        self.keyword_frame.place(x=0, y=0, relwidth=1, relheight=1)
        
        # Ключевое слово
        ttk.Label(self.keyword_frame, text="Ключевое слово:").grid(row=0, column=0, sticky=tk.W, pady=0)
        self.keyword_var_keyword = tk.StringVar(value="Строитель")
        self.keyword_entry_keyword = ttk.Entry(self.keyword_frame, textvariable=self.keyword_var_keyword, width=25)
        self.keyword_entry_keyword.grid(row=0, column=1, padx=5, pady=0, sticky=tk.W)
        
        # Город
        ttk.Label(self.keyword_frame, text="Город:").grid(row=1, column=0, sticky=tk.W, pady=0)
        self.city_var_keyword = tk.StringVar(value="Челябинск")
        self.city_entry_keyword = ttk.Entry(self.keyword_frame, textvariable=self.city_var_keyword, width=25)
        self.city_entry_keyword.grid(row=1, column=1, padx=5, pady=0, sticky=tk.W)
        
    def create_url_params(self):
        """Создание элементов для парсера по URL"""
        self.url_frame = ttk.Frame(self.params_frame)
        self.url_frame.place(x=0, y=0, relwidth=1, relheight=1)

        # URL для парсинга
        ttk.Label(self.url_frame, text="URL страницы SuperJob:").grid(row=0, column=0, sticky=tk.W, pady=0)
        self.url_var = tk.StringVar(value="https://www.superjob.ru/vacancy/search/?keywords=Москва%20Повар")
        self.url_entry = ttk.Entry(self.url_frame, textvariable=self.url_var, width=50)
        self.url_entry.grid(row=0, column=1, padx=5, pady=0, sticky=tk.W)

        # Пустое пространство для выравнивания
        empty_space = ttk.Frame(self.url_frame, height=30)
        empty_space.grid(row=1, column=0, columnspan=2, pady=0)
        
    def create_phone_params(self):
        """Создание элементов для парсера телефонов"""
        self.phone_frame = ttk.Frame(self.params_frame)
        self.phone_frame.place(x=0, y=0, relwidth=1, relheight=1)

        # Загрузить Excel файл
        ttk.Label(self.phone_frame, text="Выбрать файл:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)

        self.excel_file_btn = ttk.Button(self.phone_frame, text="Excel файл после поиска объявлений",
                                        command=self.btn_open, width=32)
        self.excel_file_btn.grid(row=0, column=1, padx=5, pady=0, sticky=tk.EW)

        self.continue_btn = ttk.Button(self.phone_frame, text="Вход выполнен", 
                                        command=self.on_continue_clicked, width=22)
        self.continue_btn.grid(row=1, column=1, padx=5, pady=0, sticky=tk.EW)


        # Путь к файлу (необязательно, но полезно)
        self.excel_file_path = tk.StringVar()
        ttk.Label(self.phone_frame, textvariable=self.excel_file_path,
                foreground="gray", wraplength=300).grid(row=0, column=2, padx=5, pady=0, sticky=tk.W)
        
    def toggle_parser_mode(self):
        """Переключение между режимами парсинга"""
        if self.parser_mode_key.get() == "keyword":
            # Показываем параметры для парсера по ключу
            self.url_frame.place_forget()
            self.phone_frame.place_forget()
            self.keyword_frame.place(x=0, y=0, relwidth=1, relheight=1)
            self.generate_url_btn.config(state=tk.NORMAL)
            self.firm_count_spinbox.config(state=tk.NORMAL)
        elif self.parser_mode_key.get() == "url":
            # Показываем параметры для парсера по URL
            self.keyword_frame.place_forget()
            self.phone_frame.place_forget()
            self.url_frame.place(x=0, y=0, relwidth=1, relheight=1)
            self.generate_url_btn.config(state=tk.DISABLED)
            self.firm_count_spinbox.config(state=tk.NORMAL)
        elif self.parser_mode_key.get() == "phone":
            # Показываем параметры для парсера телефонов
            self.keyword_frame.place_forget()
            self.url_frame.place_forget()
            self.phone_frame.place(x=0, y=0, relwidth=1, relheight=1)
            self.generate_url_btn.config(state=tk.DISABLED)
            self.firm_count_spinbox.config(state=tk.DISABLED)
    
    def run_async_parsing(self, parser_instance):
        """Запуск асинхронного парсинга в отдельном потоке"""
        try:
            # Создаем и запускаем runner
            runner = AsyncParserRunner(
                parser_instance, 
                update_callback=self.update_gui_from_thread,
                completion_callback=self.on_parsing_complete
            )
            self.parser_thread = runner.start()
            
        except Exception as e:
            self.update_gui_from_thread(f"Ошибка запуска: {str(e)}")
            self.is_parsing = False
    
    def run_parsing(self):
        """Запуск парсинга в зависимости от выбранного режима"""
        if self.is_parsing:
            messagebox.showwarning("Предупреждение", "Парсинг уже выполняется!")
            return
        self.is_parsing = True
        if self.parser_mode_key.get() == "keyword":
            self.run_keyword_parsing()
        elif self.parser_mode_key.get() == "url":
            self.run_url_parsing()
        elif self.parser_mode_key.get() == "phone":
            self.run_phone_parsing()
    
    def run_keyword_parsing(self):
        """Запуск парсинга по ключу"""
        keyword = self.keyword_var_keyword.get()
        city = self.city_var_keyword.get()
        max_vacancies = self.firm_count_var.get()

        if not keyword or not city:
            messagebox.showwarning("Предупреждение", "Заполните все поля!")
            return

        right_city = re.sub(r'[^а-яА-Яa-zA-Z\s]', '', city).strip()
        self.log_message(f"Начало парсинга по ключу: '{keyword}' в {right_city}, количество: {max_vacancies}")
        self.status_var.set(f"Парсинг по ключу: {keyword} в {city}")

        # Формируем URL для SearchSuperJob
        url = f"https://www.superjob.ru/vacancy/search/?keywords={city}%20{keyword}"

        self.is_parsing = True
        # Передаем callback для проверки флага остановки
        self.parser_instance = SearchSuperJob(
            url=url,
            max_vacancies=max_vacancies,
            stop_callback=lambda: not self.is_parsing
        )
        self.parser_thread = threading.Thread(
            target=self.run_async_parsing,
            args=(self.parser_instance,),
            daemon=True
        )
        self.parser_thread.start()
    
    def run_url_parsing(self):
        """Запуск парсинга по URL - извлекаем город и ключ из URL"""
        url = self.url_var.get()
        max_vacancies = self.firm_count_var.get()

        if not url:
            messagebox.showwarning("Предупреждение", "Введите URL для парсинга!")
            return

        # Проверяем, что это URL superjob
        if not url.startswith(('https://www.superjob.ru/', 'http://www.superjob.ru/')):
            messagebox.showwarning("Предупреждение", "Введите корректный URL superjob!")
            return

        try:
            self.log_message(f"Парсинг по URL: {url}")
            self.status_var.set(f"Парсинг по URL: {max_vacancies} вакансий")

            self.is_parsing = True
            # Передаем callback для проверки флага остановки
            self.parser_instance = SearchSuperJob(
                url=url,
                max_vacancies=max_vacancies,
                stop_callback=lambda: not self.is_parsing
            )
            runner = AsyncParserRunner(
                self.parser_instance,
                update_callback=self.update_gui_from_thread,
                completion_callback=self.on_parsing_complete
            )
            runner.start()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка парсинга по URL: {str(e)}")

    def run_phone_parsing(self):
        """Запуск парсинга телефонов из Excel файла"""
        # Проверяем, что файл выбран
        if not self.phone_excel_path or not os.path.exists(self.phone_excel_path):
            self.is_parsing = False
            self.log_message("Внимание! Сначала выберите Excel файл!")
            self.status_var.set("Выберите Excel файл!")
            return

        if not hasattr(self, 'df') or self.df is None:
            messagebox.showwarning("Предупреждение", "Файл не загружен! Выберите файл еще раз.")
            return

        # Проверяем наличие необходимой колонки
        if 'Ссылка' not in self.df.columns:
            messagebox.showerror("Ошибка",
                "В файле должна быть колонка 'Ссылка' с ссылками на объявления!")
            return

        try:
            file_name = os.path.basename(self.phone_excel_path)
            self.log_message(f"Начало парсинга телефонов из файла: {file_name}")
            self.status_var.set(f"Парсинг телефонов: {file_name}")

            self.is_parsing = True

            # Создаем экземпляр парсера телефонов с callback для проверки остановки
            self.parser_instance = VacancyParser(
                input_file=self.phone_excel_path,
                gui_mode=True,
                stop_callback=lambda: not self.is_parsing
            )

            # Запускаем парсер в отдельном потоке (синхронная версия)
            self.parser_thread = threading.Thread(
                target=self._run_sync_parser,
                args=(self.parser_instance,),
                daemon=True
            )
            self.parser_thread.start()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка запуска парсера телефонов: {str(e)}")
            self.log_message(f"Ошибка запуска парсера телефонов: {str(e)}")
            self.is_parsing = False

    def _run_sync_parser(self, parser_instance):
        """Запуск синхронного парсера в отдельном потоке"""
        try:
            self.update_gui_from_thread("Запуск парсера телефонов...")
            parser_instance.run()
            self.on_parsing_complete(flag=True)
        except Exception as e:
            self.update_gui_from_thread(f"Ошибка парсинга: {str(e)[:300]}...")
            self.on_parsing_complete(flag=False)

    def on_parsing_complete(self, flag=True):
        """Вызывается при завершении парсинга (успешном или с ошибкой)"""
        def update():
            self.is_parsing = False
            if flag:
                self.status_var.set("Парсинг успешно завершен")
                self.log_message("Парсинг успешно завершен")
            else:
                self.status_var.set("Парсинг остановлен")
                self.log_message("Парсинг остановлен")

        self.after(0, update)
        
    def stop_parsing(self):
        """Остановка парсинга"""
        if not self.is_parsing:
            self.log_message("Ничего не выполняется!")
            return

        self.is_parsing = False
        self.log_message("Остановка парсинга...")

        # Закрываем Chrome через taskkill
        import subprocess
        import os

        try:
            if os.name == 'nt':  # Windows
                result = subprocess.run(
                    ['taskkill', '/F', '/IM', 'chrome.exe', '/T'],
                    capture_output=True,
                    text=True
                )
                if result.returncode == 0:
                    self.log_message("Chrome успешно закрыт")
                else:
                    self.log_message(f"Chrome закрыт (код: {result.returncode})")
            else:  # Linux/Mac
                subprocess.run(['pkill', 'chrome'], capture_output=True)
                self.log_message("Chrome закрыт")

        except Exception as e:
            self.log_message(f"При закрытии Chrome: {str(e)}")

        self.status_var.set("Парсинг остановлен")
        self.log_message("Парсинг остановлен пользователем")
    
    def copy_ads_file_to_path(self):
        self.file_to_path(self.source_file_path)
        
    def copy_ready_file_to_path(self):
        self.file_to_path(self.output_excel)
        
    def file_to_path(self, file_path):
        """Копирование конкретного файла в выбранную папку"""
        if not os.path.exists(file_path):
            self.log_message("Ошибка экспорта объявлений! Исходный файл не найден.")
            self.status_var.set("Исходный файл не найден.")
            return
        
        target_folder = filedialog.askdirectory(
            title="Выберите папку для копирования файла"
        )
        
        if not target_folder:
            return

        try:
            filename = os.path.basename(file_path)
            target_path = os.path.join(target_folder, filename)
            
            # Проверка на существование
            if os.path.exists(target_path):
                overwrite = messagebox.askyesno(
                    "Подтверждение",
                    f"Файл '{filename}' уже существует. Заменить?"
                )
                if not overwrite:
                    return
            
            shutil.copy2(file_path, target_path)
            
            self.log_message(f"Успех! Файл '{filename}' успешно скопирован в:\n{target_folder}")
            self.status_var.set(f"Файл '{filename}' успешно скопирован!")
            
        except Exception as e:
            self.log_message(f"Ошибка! Не удалось скопировать файл:\n{str(e)}")
            self.status_var.set("Не удалось скопировать файл.")
        
    def create_status_bar(self):
        """Создание строки состояния"""
        self.status_var = tk.StringVar()
        self.status_var.set("Готов к работе")
        self.status_bar = ttk.Label(self, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, padding=(10, 5))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def load_data(self, file_path):
        """Загружаю Excel файл"""
        try:
            df = pd.read_excel(file_path, na_values=["--.--", "nan", "NaN", "", "---"])
            return df
        except Exception as e:
            print(f"Ошибка загрузки данных: {e}")
            raise
            
    def btn_open(self):
        """Обработчик кнопки 'Excel файл после поиска обновлений'"""
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл с ссылками на объявления",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if file_path:
            # Сохраняем путь для парсера телефонов
            self.phone_excel_path = file_path
            
            # Отображаем имя файла в интерфейсе
            file_name = os.path.basename(file_path)
            if len(file_name) > 19:
                # Обрезаем первые 20 символов, добавляем "...", затем пробел и расширение
                file_basename = file_name[:14] + "... " + file_name[file_name.rfind('.'):]
            else:
                file_basename = file_name
            self.excel_file_path.set(f"Выбран: {file_basename}")

            self.update_idletasks()
            try:
                # Загружаем для проверки
                self.df = self.load_data(file_path)
                
                # Проверяем наличие необходимой колонки
                if 'Ссылка' not in self.df.columns:
                    messagebox.showwarning("Предупреждение", 
                        "В файле должна быть колонка 'Ссылка'!")
                    self.phone_excel_path = None
                else:
                    self.log_message(f"Excel файл успешно загружен!")
                    self.log_message(f"Количество объявлений: {len(self.df)}")
                    self.log_message(f"Теперь можете запустить парсинг телефонов.")
                    self.status_var.set(f"Количество объявлений в Excel: {len(self.df)}")
            
                    
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")
                self.status_var.set("Ошибка загрузки файла")
                self.phone_excel_path = None

    def on_continue_clicked(self):
        """Обработчик нажатия кнопки 'Вход выполнен'"""
        try:
            if hasattr(self, 'parser_instance') and self.parser_instance:
                # Отправляем подтверждение в парсер
                self.parser_instance.trigger_enter_from_gui()
                self.log_message("Подтверждение входа отправлено парсеру")
                self.status_var.set("Парсинг продолжается...")
            else:
                self.log_message("Ошибка: парсер не инициализирован")
        except Exception as e:
            self.log_message(f"Ошибка отправки подтверждения: {str(e)}")


    def open_link(self):
        webbrowser.open("https://github.com/itrickon/SJobParser") 
        
    def hotkeys_info(self):
        """Обработчик кнопки 'Горячие клавиши'"""
        # Создаем собственное окно вместо messagebox
        top = Toplevel()
        top.title("Горячие клавиши")
        
        # Создаем Frame для размещения текстового виджета и скроллбара
        frame = tk.Frame(top)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Создаем текстовое поле
        text_widget = Text(frame, wrap=tk.WORD, width=60, height=12, 
                        font=("Arial", 10))
        
        
        top.resizable(False, False)
        
        # Добавляем остальной текст
        cities = [
        "       Горячие клавиши приложения:\n",
        "   Основные операции:\n",
        "     • Ctrl + O   - Открыть Excel файл...\n",
        "     • Ctrl + R   - Запустить парсинг\n",
        "     • Ctrl + S   - Остановить парсинг\n",
        "     • Ctrl + L   - Очистить лог\n",
        "     • Ctrl + Q   - Выйти из приложения\n",
        "   Дополнительные:\n",
        "     • Ctrl + G - Сгенерировать URL (в режиме по ключу)\n",
        "     • F1         - Руководство пользователя\n",
        "   Сочетания клавиш работают в любом месте приложения.\n",
        ]
        
        for city_text in cities:
            text_widget.insert(tk.END, city_text)
        
        text_widget.configure(state='disabled')  # Только для чтения
        
        # Кнопка закрытия
        button = tk.Button(top, text="Закрыть", command=top.destroy)
        
        text_widget.pack()
        button.pack(pady=10)
        
        # Центрируем окно
        top.update_idletasks()
        width = top.winfo_width()
        height = top.winfo_height()
        x = (top.winfo_screenwidth() // 2) - (width // 2)
        y = (top.winfo_screenheight() // 2) - (height // 2)
        top.geometry(f'{width}x{height}+{x}+{y}')

    def btn_about(self):
        """Обработчик кнопки 'О программе'"""
        # Создаем собственное окно вместо messagebox
        top = Toplevel()
        top.title("О программе")
        
        # Создаем Frame для размещения текстового виджета и скроллбара
        frame = tk.Frame(top)
        frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Создаем текстовое поле
        text_widget = Text(frame, wrap=tk.WORD, width=67, height=20, 
                        font=("Arial", 10))
        
        top.resizable(False, False)
        
        # Добавляем остальной текст
        about_text = [
        "       SJobParser\n\n",
        "  Данный инструмент предназначен для сбора открытой информации в образовательных и исследовательских целях.\n\n",
        "    Версия 0.1.1\n\n",
        "  Режимы работы:\n",
        "    1. Парсер по ключу - поиск организаций по ключевому слову и городу\n",
        "    2. Парсер по URL - парсинг конкретной страницы поиска\n\n",
        "  Возможности:\n",
        "    • Поддержка светлой и темной темы\n\n",
        "  Используемые технологии:\n",
        "    • Python 3.11+\n",
        "    • Playwright для веб-скрапинга\n",
        "    • tkinter для графического интерфейса\n",
        "    • sv_ttk для современных стилей\n\n",
        "    https://github.com/itrickon/SJobParser",
        ]
        
        for city_text in about_text:
            text_widget.insert(tk.END, city_text)
        
        text_widget.configure(state='disabled')  # Только для чтения
        
        # Кнопка закрытия
        button = tk.Button(top, text="Закрыть", command=top.destroy)
        
        text_widget.pack()
        button.pack(pady=10)
        
        # Центрируем окно
        top.update_idletasks()
        width = top.winfo_width()
        height = top.winfo_height()
        x = (top.winfo_screenwidth() // 2) - (width // 2)
        y = (top.winfo_screenheight() // 2) - (height // 2)
        top.geometry(f'{width}x{height}+{x}+{y}')
 
    def log_message(self, message):
        """Добавление сообщения в лог с цветами"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        
        # Определяем уровень
        msg_lower = message.lower()
        error_words = ["ошибка", "error", "closed", "exception", "failed", "прервано"]
        warning_words = ["предупреждение", "warning", "внимание", "остановлен"]
        success_words = ["успешно", "success", "завершен", "готово", "успешн"]
        
        if any(word in msg_lower for word in error_words):
            level = "ERROR"
        elif any(word in msg_lower for word in warning_words):
            level = "WARNING"
        elif any(word in msg_lower for word in success_words):
            level = "SUCCESS"
        else:
            level = "INFO"
        
        formatted_message = f"[{timestamp}] [{level}] {message}\n"
        
        # Вставляем с тегом
        self.log_text.insert(tk.END, formatted_message, (level,))
        self.log_text.see(tk.END)

    def clear_log(self):
        """Очистка лога"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("Лог очищен")
        self.status_var.set("Лог очищен")
 
    def update_gui_from_thread(self, message):
        """Обновление GUI из потока"""
        def update():
            self.log_message(message)
            self.status_var.set(message[:50] + "..." if len(message) > 50 else message)
            
        self.after(0, update)
 
    def btn_exit(self):
        """Выход из приложения"""
        if self.is_parsing:
            if not messagebox.askyesno("Предупреждение", 
                                      "Парсинг выполняется. Вы уверены, что хотите выйти?"):
                return
        
        if messagebox.askyesno("Выход", "Вы уверены, что хотите выйти?"):
            if self.is_parsing:
                self.stop_parsing()
            self.parent.quit()
        
def main():
    """Точка входа в приложение"""
    root = tk.Tk()
    app = SJobParser(root)
    root.mainloop()


if __name__ == "__main__":
    main()