import os
import time
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import logging
import pandas as pd
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
import re
import json
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('parser.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)


class YandexImagesLinkParser:
    def __init__(self, output_file="image_links.xlsx", file_format='excel'):
        """
        Инициализация парсера
        
        Args:
            output_file: имя выходного файла
            file_format: 'excel' или 'csv'
        """
        self.driver = None
        self.wait = None
        self.output_file = output_file
        self.file_format = file_format.lower()
        self.processed_count = 0
        self.success_count = 0
        self.logger = logging.getLogger(__name__)
        self.results = []
        
        # Проверяем соответствие расширения и формата
        self._validate_file_format()
        
        # Создаем файл с заголовками если его нет
        if not os.path.exists(output_file):
            self._create_empty_file()
    
    def _validate_file_format(self):
        """Проверка соответствия расширения файла и формата"""
        ext = os.path.splitext(self.output_file)[1].lower()
        
        if self.file_format == 'excel' and ext not in ['.xlsx', '.xls']:
            self.logger.warning(f"⚠️ Для формата Excel рекомендуется расширение .xlsx или .xls. Текущее: {ext}")
            # Автоматически исправляем расширение если нужно
            if ext == '.csv':
                self.output_file = self.output_file.replace('.csv', '.xlsx')
                self.logger.info(f"🔄 Автоматически изменено расширение на: {self.output_file}")
        
        elif self.file_format == 'csv' and ext != '.csv':
            self.logger.warning(f"⚠️ Для формата CSV рекомендуется расширение .csv. Текущее: {ext}")
            if ext in ['.xlsx', '.xls']:
                self.output_file = self.output_file.replace('.xlsx', '.csv').replace('.xls', '.csv')
                self.logger.info(f"🔄 Автоматически изменено расширение на: {self.output_file}")
    
    def _create_empty_file(self):
        """Создание пустого файла с заголовками"""
        try:
            if self.file_format == 'excel':
                df = pd.DataFrame(columns=['query', 'image_url', 'timestamp'])
                df.to_excel(self.output_file, index=False, engine='openpyxl')
                self.logger.info(f"✅ Создан новый Excel файл: {self.output_file} с заголовками")
            else:
                df = pd.DataFrame(columns=['query', 'image_url', 'timestamp'])
                df.to_csv(self.output_file, index=False, encoding='utf-8-sig')
                self.logger.info(f"✅ Создан новый CSV файл: {self.output_file} с заголовками")
        except Exception as e:
            self.logger.error(f"❌ Ошибка создания файла: {str(e)}")
            # Пробуем создать CSV как резерв
            try:
                csv_file = self.output_file.replace('.xlsx', '.csv').replace('.xls', '.csv')
                df = pd.DataFrame(columns=['query', 'image_url', 'timestamp'])
                df.to_csv(csv_file, index=False, encoding='utf-8-sig')
                self.logger.info(f"✅ Создан резервный CSV файл: {csv_file}")
                self.output_file = csv_file
                self.file_format = 'csv'
            except:
                pass
    
    def get_random_user_agent(self):
        """Случайный User-Agent"""
        user_agents = [
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/121.0'
        ]
        return random.choice(user_agents)
    
    def setup_driver(self):
        """Настройка Chrome драйвера"""
        try:
            chrome_options = Options()
            
            # Основные настройки
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            chrome_options.add_argument('--disable-gpu')
            chrome_options.add_argument('--no-sandbox')
            chrome_options.add_argument('--disable-dev-shm-usage')
            chrome_options.add_argument('--window-size=1920,1080')
            chrome_options.add_argument(f'--user-agent={self.get_random_user_agent()}')
            
            # Дополнительные настройки для обхода
            chrome_options.add_argument('--disable-web-security')
            chrome_options.add_argument('--disable-features=IsolateOrigins,site-per-process')
            chrome_options.add_argument('--disable-site-isolation-trials')
            chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            
            # Отключаем автоматизацию
            chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            chrome_options.add_experimental_option('useAutomationExtension', False)
            
            # Настройки для эмуляции реального браузера
            chrome_options.add_experimental_option("prefs", {
                "credentials_enable_service": False,
                "profile.password_manager_enabled": False,
                "profile.default_content_setting_values.notifications": 2,
                "profile.default_content_settings.popups": 0,
            })
            
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            
            # Удаляем признаки автоматизации
            self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            self.driver.execute_script("Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]})")
            
            self.wait = WebDriverWait(self.driver, 15)
            self.logger.info("✅ Драйвер успешно настроен")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Ошибка при настройке драйвера: {str(e)}")
            return False
    
    def random_delay(self, min_sec=1, max_sec=3):
        """Случайная задержка"""
        time.sleep(random.uniform(min_sec, max_sec))
    
    def human_like_scroll(self):
        """Эмуляция скролла как у человека"""
        try:
            # Плавный скролл вниз
            for _ in range(random.randint(1, 3)):
                scroll_amount = random.randint(200, 600)
                self.driver.execute_script(f"window.scrollBy(0, {scroll_amount})")
                time.sleep(random.uniform(0.3, 0.8))
            
            # Иногда скроллим вверх
            if random.random() < 0.3:
                self.driver.execute_script(f"window.scrollBy(0, -{random.randint(100, 300)})")
                time.sleep(random.uniform(0.2, 0.5))
                
        except Exception as e:
            self.logger.debug(f"Ошибка скролла: {e}")
    
    def extract_image_url_from_element(self, element):
        """Извлечение URL изображения из элемента"""
        try:
            # Пробуем разные атрибуты
            url = element.get_attribute('src')
            if not url or url.startswith('data:image'):
                url = element.get_attribute('data-src')
            if not url or url.startswith('data:image'):
                url = element.get_attribute('data-url')
            if not url or url.startswith('data:image'):
                url = element.get_attribute('data-original')
            if not url or url.startswith('data:image'):
                url = element.get_attribute('content')
            
            # Если URL относительный, делаем абсолютным
            if url and url.startswith('//'):
                url = 'https:' + url
            elif url and url.startswith('/'):
                url = 'https://yandex.ru' + url
            
            # Фильтруем мусор
            if url and not url.startswith('data:image') and len(url) > 10:
                return url
            
            return None
            
        except Exception as e:
            return None
    
    def get_first_image_link(self, query):
        """Получение ссылки на первое изображение по запросу"""
        try:
            # Очищаем запрос
            clean_query = query.replace('/', ' ').replace('\\', ' ').replace('*', ' ').replace('?', ' ')
            search_url = f"https://yandex.ru/images/search?text={clean_query.replace(' ', '+')}"
            
            self.logger.info(f"🔍 Поиск: {clean_query[:50]}...")
            
            # Переходим на страницу поиска
            self.driver.get(search_url)
            self.random_delay(2, 4)
            
            # Эмулируем поведение человека
            self.human_like_scroll()
            self.random_delay(1, 2)
            
            # Ждем загрузки результатов
            try:
                # Пробуем разные селекторы для поиска изображений
                image_selectors = [
                    "img.serp-item__thumb",
                    "img[class*='serp-item']",
                    "div.serp-item img",
                    "a.serp-item__link img",
                    "img[src*='yandex.net']",
                    "img[data-src*='yandex.net']",
                    "div[class*='image'] img",
                    "img[class*='image']"
                ]
                
                image_url = None
                image_element = None
                
                # Ищем первое изображение
                for selector in image_selectors:
                    try:
                        elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        for elem in elements:
                            # Проверяем, что элемент видимый
                            if elem.is_displayed():
                                url = self.extract_image_url_from_element(elem)
                                if url:
                                    image_url = url
                                    image_element = elem
                                    self.logger.info(f"✅ Найдено изображение по селектору: {selector}")
                                    break
                        if image_url:
                            break
                    except Exception as e:
                        continue
                
                # Если не нашли через селекторы, ищем через JavaScript
                if not image_url:
                    self.logger.info("🔍 Пробуем найти через JavaScript...")
                    try:
                        image_url = self.driver.execute_script("""
                            // Ищем все изображения на странице
                            const images = document.querySelectorAll('img');
                            for (let img of images) {
                                // Проверяем видимость
                                const rect = img.getBoundingClientRect();
                                if (rect.width > 50 && rect.height > 50) {
                                    // Проверяем src
                                    let src = img.src || img.getAttribute('data-src') || img.getAttribute('data-url');
                                    if (src && !src.startsWith('data:image') && src.includes('yandex.net')) {
                                        return src;
                                    }
                                }
                            }
                            return null;
                        """)
                        if image_url:
                            self.logger.info("✅ Найдено через JavaScript")
                    except Exception as e:
                        pass
                
                # Если нашли URL, проверяем его
                if image_url:
                    # Нормализуем URL
                    if image_url.startswith('//'):
                        image_url = 'https:' + image_url
                    
                    # Проверяем, что URL валидный
                    if image_url.startswith('http') and not image_url.startswith('data:'):
                        self.logger.info(f"✅ Ссылка получена: {image_url[:100]}...")
                        return image_url
                    else:
                        self.logger.warning(f"⚠️ Невалидная ссылка: {image_url[:50]}")
                        return None
                
                # Если не нашли изображение, пытаемся кликнуть на результат
                self.logger.info("🔍 Пробуем кликнуть на результат...")
                try:
                    # Ищем первый результат
                    result_selectors = [
                        "div.serp-item a",
                        "a.serp-item__link",
                        "div[class*='serp-item'] a"
                    ]
                    
                    for selector in result_selectors:
                        try:
                            elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                            for elem in elements:
                                if elem.is_displayed():
                                    # Кликаем на результат
                                    self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elem)
                                    self.random_delay(0.5, 1)
                                    self.driver.execute_script("arguments[0].click();", elem)
                                    self.random_delay(2, 3)
                                    
                                    # После клика ищем изображение на открывшейся странице
                                    image_url = self.get_image_url_from_opened_page()
                                    if image_url:
                                        return image_url
                                    break
                            if image_url:
                                break
                        except:
                            continue
                            
                except Exception as e:
                    self.logger.debug(f"Ошибка при клике: {e}")
                
                return None
                
            except TimeoutException:
                self.logger.warning("⏰ Таймаут загрузки страницы")
                return None
            except Exception as e:
                self.logger.warning(f"⚠️ Ошибка при получении ссылки: {str(e)}")
                return None
                
        except Exception as e:
            self.logger.error(f"❌ Ошибка поиска: {str(e)}")
            return None
    
    def get_image_url_from_opened_page(self):
        """Получение URL изображения после клика на результат"""
        try:
            # Ждем загрузки изображения
            self.random_delay(1, 2)
            
            # Ищем изображение в открытом просмотрщике
            selectors = [
                "img.MMImage-Origin",
                "img.MMImage",
                "div.MMImageContainer img",
                "img[class*='origin']",
                "img[src*='yandex.net']:not([src*='data:image'])",
                "div.image-view img",
                "img.image__image"
            ]
            
            for selector in selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elements:
                        if elem.is_displayed():
                            url = self.extract_image_url_from_element(elem)
                            if url:
                                self.logger.info(f"✅ Найдено в просмотрщике")
                                return url
                except:
                    continue
            
            # Пробуем получить из URL страницы
            try:
                current_url = self.driver.current_url
                if 'img_url' in current_url:
                    parsed = urlparse(current_url)
                    params = parse_qs(parsed.query)
                    if 'img_url' in params:
                        return params['img_url'][0]
            except:
                pass
            
            return None
            
        except Exception as e:
            self.logger.debug(f"Ошибка получения из просмотрщика: {e}")
            return None
    
    def _save_dataframe(self, df):
        """Универсальное сохранение DataFrame в файл"""
        try:
            if self.file_format == 'excel':
                df.to_excel(self.output_file, index=False, engine='openpyxl')
            else:
                df.to_csv(self.output_file, index=False, encoding='utf-8-sig')
            return True
        except Exception as e:
            self.logger.error(f"❌ Ошибка сохранения: {str(e)}")
            return False
    
    def _read_existing_data(self):
        """Чтение существующих данных из файла"""
        try:
            if not os.path.exists(self.output_file):
                return None
            
            if self.file_format == 'excel':
                return pd.read_excel(self.output_file, engine='openpyxl')
            else:
                return pd.read_csv(self.output_file, encoding='utf-8-sig')
        except Exception as e:
            self.logger.warning(f"⚠️ Ошибка чтения файла: {e}")
            return None
    
    def save_result_to_excel(self, result):
        """
        Сохранение результата с раздельными колонками:
        - query (запрос)
        - image_url (ссылка на изображение)
        - timestamp (время)
        """
        try:
            # Создаем DataFrame с новой записью
            new_row = pd.DataFrame([{
                'query': result['query'],
                'image_url': result['image_url'],
                'timestamp': result['timestamp']
            }])
            
            # Проверяем, существует ли файл
            if os.path.exists(self.output_file):
                try:
                    # Читаем существующий файл
                    existing_df = self._read_existing_data()
                    
                    if existing_df is not None and not existing_df.empty:
                        # Проверяем, есть ли нужные колонки
                        required_columns = ['query', 'image_url', 'timestamp']
                        for col in required_columns:
                            if col not in existing_df.columns:
                                existing_df[col] = None
                        
                        # Объединяем
                        combined_df = pd.concat([existing_df, new_row], ignore_index=True)
                    else:
                        combined_df = new_row
                    
                    # Сохраняем
                    if self._save_dataframe(combined_df):
                        self.logger.debug(f"💾 Добавлена строка в {self.output_file}")
                    
                except Exception as e:
                    self.logger.warning(f"⚠️ Ошибка чтения файла: {e}. Создаем новый.")
                    if self._save_dataframe(new_row):
                        self.logger.info(f"✅ Создан новый файл: {self.output_file}")
            else:
                # Создаем новый файл с первой записью
                if self._save_dataframe(new_row):
                    self.logger.info(f"✅ Создан новый файл: {self.output_file}")
                
        except Exception as e:
            self.logger.error(f"❌ Ошибка сохранения: {str(e)}")
            
            # Резервное сохранение в CSV
            try:
                backup_file = self.output_file
                if self.file_format == 'excel':
                    backup_file = self.output_file.replace('.xlsx', '_backup.csv').replace('.xls', '_backup.csv')
                else:
                    backup_file = self.output_file.replace('.csv', '_backup.csv')
                
                df = pd.DataFrame([{
                    'query': result['query'],
                    'image_url': result['image_url'],
                    'timestamp': result['timestamp']
                }])
                
                if os.path.exists(backup_file):
                    existing_df = pd.read_csv(backup_file, encoding='utf-8-sig')
                    df = pd.concat([existing_df, df], ignore_index=True)
                
                df.to_csv(backup_file, index=False, encoding='utf-8-sig')
                self.logger.info(f"💾 Сохранено в резервный CSV: {backup_file}")
            except Exception as e2:
                self.logger.error(f"❌ Критическая ошибка сохранения: {e2}")
    
    def process_query(self, query):
        """Обработка одного запроса - получение ссылки на изображение"""
        if not query or not str(query).strip():
            return None
        
        query = str(query).strip()
        self.processed_count += 1
        
        if len(query) > 200:
            self.logger.warning(f"⚠️ Запрос слишком длинный: {query[:50]}...")
            return None
        
        self.logger.info(f"[{self.processed_count}] {query[:50]}")
        
        # Случайная задержка перед запросом
        self.random_delay(2, 5)
        
        try:
            image_url = self.get_first_image_link(query)
            
            if image_url:
                self.success_count += 1
                # Сохраняем результат в список
                result = {
                    'query': query,
                    'image_url': image_url,
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                self.results.append(result)
                
                # Сохраняем сразу в файл с раздельными колонками
                self.save_result_to_excel(result)
                
                self.logger.info(f"✅ Ссылка получена: {image_url[:80]}...")
                return image_url
            else:
                self.logger.warning(f"❌ Ссылка не найдена: {query[:50]}")
                return None
                
        except Exception as e:
            self.logger.error(f"❌ Ошибка обработки: {str(e)}")
            return None
    
    def parse_excel(self, excel_file, column_name=None, limit=None, start_from=0):
        """Чтение Excel или CSV файла"""
        try:
            self.logger.info(f"📂 Чтение файла: {os.path.basename(excel_file)}")
            
            if not os.path.exists(excel_file):
                self.logger.error(f"❌ Файл не найден: {excel_file}")
                return []
            
            # Определяем формат файла по расширению
            ext = os.path.splitext(excel_file)[1].lower()
            
            try:
                if ext in ['.xlsx', '.xls']:
                    df = pd.read_excel(excel_file, engine='openpyxl')
                else:
                    df = pd.read_csv(excel_file, encoding='utf-8-sig')
            except Exception as e:
                self.logger.error(f"❌ Ошибка чтения: {str(e)}")
                return []
            
            self.logger.info(f"📊 Найдено строк: {len(df)}, столбцов: {len(df.columns)}")
            
            if column_name and column_name in df.columns:
                items = df[column_name].dropna().tolist()
            else:
                # Ищем текстовый столбец
                text_cols = []
                for col in df.columns:
                    if df[col].dtype == 'object':
                        non_empty = df[col].dropna()
                        if len(non_empty) > 0 and isinstance(non_empty.iloc[0], str):
                            text_cols.append(col)
                
                if text_cols:
                    first_col = text_cols[0]
                    self.logger.info(f"📌 Используем столбец: {first_col}")
                    items = df[first_col].dropna().tolist()
                else:
                    first_col = df.columns[0]
                    self.logger.info(f"📌 Используем столбец: {first_col}")
                    items = df[first_col].dropna().tolist()
            
            items = [str(item).strip() for item in items if str(item).strip()]
            
            if start_from > 0:
                items = items[start_from:]
            
            if limit and limit > 0:
                items = items[:limit]
            
            self.logger.info(f"✅ Загружено {len(items)} позиций")
            return items
            
        except Exception as e:
            self.logger.error(f"❌ Ошибка: {str(e)}")
            return []
    
    def process_list(self, items, batch_size=30):
        """Обработка списка запросов"""
        if not items:
            self.logger.warning("⚠️ Список пуст")
            return {
                'total': 0, 
                'success': 0, 
                'failed': 0, 
                'results': self.results
            }
        
        self.processed_count = 0
        self.success_count = 0
        self.results = []
        
        # Загружаем существующие результаты если есть
        existing_queries = []
        if os.path.exists(self.output_file):
            try:
                existing_df = self._read_existing_data()
                if existing_df is not None and 'query' in existing_df.columns:
                    existing_queries = existing_df['query'].tolist()
                    self.logger.info(f"📋 Найдено {len(existing_queries)} ранее обработанных запросов")
            except Exception as e:
                self.logger.warning(f"⚠️ Не удалось прочитать существующий файл: {e}")
                existing_queries = []
        
        results = {
            'total': len(items),
            'success': 0,
            'failed': 0,
            'results': self.results
        }
        
        progress_file = "progress.txt"
        
        for i, item in enumerate(items, 1):
            try:
                # Пропускаем уже обработанные
                if item in existing_queries:
                    self.logger.info(f"⏭ Пропуск (уже обработано): {item[:50]}")
                    continue
                
                self.processed_count = i
                image_url = self.process_query(item)
                
                if image_url:
                    results['success'] += 1
                else:
                    results['failed'] += 1
                
                # Сохраняем прогресс
                if i % 5 == 0:
                    with open(progress_file, 'w', encoding='utf-8') as f:
                        f.write(f"Обработано: {i}/{len(items)}\n")
                        f.write(f"Успешно: {results['success']}\n")
                        f.write(f"Ошибок: {results['failed']}\n")
                        f.write(f"Всего ссылок: {len(self.results)}\n")
                        if i > 0:
                            f.write(f"Успешность: {results['success']/i*100:.1f}%\n")
                
                # Задержка между запросами
                if i < len(items):
                    delay = random.uniform(3, 6)
                    self.logger.info(f"⏳ Пауза {delay:.1f} сек...")
                    time.sleep(delay)
                    
            except Exception as e:
                self.logger.error(f"❌ Ошибка обработки {item}: {str(e)}")
                results['failed'] += 1
            
            # Перезапуск драйвера каждые batch_size запросов
            if i % batch_size == 0 and i > 0:
                self.logger.info("🔄 Перезапуск драйвера...")
                self.close()
                time.sleep(5)
                if not self.setup_driver():
                    self.logger.error("❌ Не удалось перезапустить драйвер")
                    break
                time.sleep(3)
        
        # Итоговый отчет
        with open(progress_file, 'w', encoding='utf-8') as f:
            f.write(f"✅ ЗАВЕРШЕНО!\n")
            f.write(f"Обработано: {len(items)}/{len(items)}\n")
            f.write(f"Успешно: {results['success']}\n")
            f.write(f"Ошибок: {results['failed']}\n")
            f.write(f"Всего ссылок: {len(self.results)}\n")
            if len(items) > 0:
                f.write(f"Успешность: {results['success']/len(items)*100:.1f}%\n")
        
        results['results'] = self.results
        self.logger.info(f"✅ Завершено! Успешно: {results['success']}, Ошибок: {results['failed']}")
        self.logger.info(f"📁 Результаты сохранены в: {self.output_file}")
        
        return results
    
    def close(self):
        """Закрытие драйвера"""
        if self.driver:
            try:
                self.driver.quit()
                self.logger.info("🔒 Драйвер закрыт")
            except:
                pass


# Пример использования
if __name__ == "__main__":
    # Пример 1: Сохранение в Excel (по умолчанию)
    print("📊 Пример 1: Сохранение в Excel")
    parser_excel = YandexImagesLinkParser(
        output_file="image_links.xlsx",
        file_format='excel'
    )
    
    try:
        if not parser_excel.setup_driver():
            print("❌ Не удалось настроить драйвер")
            exit(1)
        
        queries = [
            "красивый закат",
            "горный пейзаж",
            "цветущий сад",
            "морской пейзаж",
            "городской пейзаж"
        ]
        
        results = parser_excel.process_list(queries, batch_size=30)
        
        print(f"\n📊 Результаты (Excel):")
        print(f"  Всего запросов: {results['total']}")
        print(f"  Успешно: {results['success']}")
        print(f"  Ошибок: {results['failed']}")
        print(f"  Ссылок сохранено: {len(results['results'])}")
        print(f"  Файл: {parser_excel.output_file}")
        
        if results['results']:
            print(f"\n📋 Первые 3 ссылки:")
            for i, result in enumerate(results['results'][:3], 1):
                print(f"  {i}. Запрос: {result['query'][:30]}...")
                print(f"     Ссылка: {result['image_url'][:80]}...")
                print(f"     Время: {result['timestamp']}")
                print()
        
    except KeyboardInterrupt:
        print("\n⚠️ Прервано пользователем")
    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
    finally:
        parser_excel.close()
    
    print("\n" + "="*50 + "\n")
    
    # Пример 2: Сохранение в CSV
    print("📊 Пример 2: Сохранение в CSV")
    parser_csv = YandexImagesLinkParser(
        output_file="image_links.csv",
        file_format='csv'
    )
    
    try:
        if not parser_csv.setup_driver():
            print("❌ Не удалось настроить драйвер")
            exit(1)
        
        queries_small = [
            "ночной город",
            "звездное небо"
        ]
        
        results = parser_csv.process_list(queries_small, batch_size=30)
        
        print(f"\n📊 Результаты (CSV):")
        print(f"  Всего запросов: {results['total']}")
        print(f"  Успешно: {results['success']}")
        print(f"  Ошибок: {results['failed']}")
        print(f"  Ссылок сохранено: {len(results['results'])}")
        print(f"  Файл: {parser_csv.output_file}")
        
    except KeyboardInterrupt:
        print("\n⚠️ Прервано пользователем")
    except Exception as e:
        print(f"❌ Критическая ошибка: {e}")
    finally:
        parser_csv.close()