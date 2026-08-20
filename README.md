Yandex Images Parser & Downloader
<p align="center"> <img src="https://img.shields.io/badge/Python-3.8%2B-blue" alt="Python Version"> <img src="https://img.shields.io/badge/Selenium-4.0%2B-green" alt="Selenium"> <img src="https://img.shields.io/badge/License-MIT-yellow" alt="License"> </p>
📖 Описание проекта
Комплексный инструмент для автоматического сбора и скачивания изображений с Яндекс.Картинок. Проект состоит из двух модулей:

Парсер ссылок (yandex_parser.py) - собирает прямые ссылки на изображения по текстовым запросам

Загрузчик изображений (image_downloader/) - массово скачивает файлы по полученным ссылкам

⚠️ Важно: Проект предназначен исключительно для образовательных и исследовательских целей. Соблюдайте авторские права и условия использования сервиса Яндекс.Картинки.

🚀 Возможности
Парсер ссылок
✅ Автоматический сбор ссылок на изображения по текстовым запросам

✅ Поддержка Excel и CSV файлов для массовой обработки

✅ Эмуляция человеческого поведения для обхода антибот-систем

✅ Автоматическое сохранение прогресса

✅ Возобновление работы с последней обработанной позиции

✅ Логирование всех действий

Загрузчик изображений
✅ Многопоточная загрузка (до 10 потоков одновременно)

✅ Автоматическое определение имени файла

✅ Пропуск уже существующих файлов

✅ Повторные попытки при ошибках

✅ Поддержка Excel, CSV и текстовых файлов

✅ Прогресс-бар загрузки

✅ Сохранение списка неудачных загрузок

📋 Требования
Python 3.8 или выше

Google Chrome браузер (последняя версия)

Доступ к интернету

🔧 Установка
1. Клонирование репозитория
bash
git clone https://github.com/yourusername/yandex-images-parser.git
cd yandex-images-parser
2. Установка зависимостей
Установите все необходимые библиотеки одной командой:

bash
pip install -r requirements.txt
Или установите вручную:

bash
# Основные зависимости
pip install selenium==4.15.0
pip install webdriver-manager==4.0.1
pip install pandas==2.1.3
pip install openpyxl==3.1.2
pip install requests==2.31.0
pip install pillow==10.1.0

# Дополнительные (опционально)
pip install undetected-chromedriver==3.5.4  # Для улучшенного обхода защиты
pip install fake-useragent==1.4.0           # Для случайных User-Agent
Содержимое requirements.txt:

txt
selenium>=4.15.0
webdriver-manager>=4.0.1
pandas>=2.1.3
openpyxl>=3.1.2
requests>=2.31.0
Pillow>=10.1.0
undetected-chromedriver>=3.5.4
fake-useragent>=1.4.0
📁 Структура проекта
text
yandex-images-parser/
├── README.md                          # Этот файл
├── requirements.txt                   # Зависимости
├── yandex_parser.py                   # Парсер ссылок на изображения
├── image_downloader/                  # Папка с загрузчиком
│   ├── __init__.py
│   ├── main.py                        # Основной скрипт загрузчика
│   ├── batch_downloader.py            # Класс загрузчика
│   └── run_downloader.py              # CLI интерфейс
├── downloads/                         # Папка для скачанных изображений (создается автоматически)
├── image_links.xlsx                   # Результат работы парсера
├── image_links.csv                    # Альтернативный формат результата
├── progress.txt                       # Файл прогресса
├── parser.log                         # Лог работы парсера
├── failed_queries.txt                 # Неудачные запросы
└── examples/                          # Примеры файлов
    ├── queries.xlsx                   # Пример Excel с запросами
    └── queries.txt                    # Пример текстового файла
🎯 Инструкция по запуску
Шаг 1: Сбор ссылок на изображения
Запустите парсер для сбора ссылок:

bash
# Базовый запуск (использует встроенные запросы)
python yandex_parser.py

# Запуск с указанием своего файла
python yandex_parser.py --input queries.xlsx --column Query --output image_links.xlsx

# Запуск с ограничением
python yandex_parser.py --input queries.xlsx --limit 50 --start 10
Аргументы командной строки:

Аргумент	Описание	По умолчанию
--input	Путь к Excel/CSV файлу с запросами	-
--column	Название столбца с запросами	-
--output	Имя выходного файла	image_links.xlsx
--format	Формат выхода (excel или csv)	excel
--limit	Ограничить количество запросов	∞
--start	Начать с N-ой строки	0
Пример использования парсера в коде:

python
from yandex_parser import YandexImagesLinkParser

# Создаем парсер
parser = YandexImagesLinkParser(
    output_file="image_links.xlsx",
    file_format='excel'  # или 'csv'
)

# Настраиваем драйвер
if parser.setup_driver():
    # Список запросов
    queries = [
        "красивый закат",
        "горный пейзаж",
        "цветущий сад"
    ]
    
    # Обрабатываем запросы
    results = parser.process_list(queries, batch_size=30)
    
    print(f"✅ Успешно: {results['success']}")
    print(f"❌ Ошибок: {results['failed']}")

# Закрываем драйвер
parser.close()
Результат работы парсера - файл image_links.xlsx с колонками:

query - текстовый запрос

image_url - прямая ссылка на изображение

timestamp - дата и время получения ссылки

Шаг 2: Массовое скачивание изображений
Перейдите в папку загрузчика и запустите скачивание:

bash
cd image_downloader

# Загрузка из Excel файла
python run_downloader.py ../image_links.xlsx --url-col image_url --output ../downloads

# Загрузка из CSV файла
python run_downloader.py ../image_links.csv --url-col image_url --output ../downloads

# Загрузка из текстового файла
python run_downloader.py ../urls.txt --output ../downloads

# Загрузка с дополнительными параметрами
python run_downloader.py ../image_links.xlsx \
    --url-col image_url \
    --name-col query \
    --output ../downloads \
    --workers 5 \
    --limit 100
Аргументы командной строки загрузчика:

Аргумент	Описание	По умолчанию
source	Путь к файлу с ссылками	Обязательный
--url-col	Название столбца с URL	url
--name-col	Название столбца с именами файлов	-
--output	Папка для сохранения	downloads
--workers	Количество параллельных загрузок	3
--limit	Ограничить количество файлов	∞
--subfolder	Подпапка внутри output	-
--timeout	Таймаут загрузки (сек)	60
--retry	Количество попыток	3
Пример использования загрузчика в коде:

python
from image_downloader.batch_downloader import BatchFileDownloader

# Создаем загрузчик
downloader = BatchFileDownloader(
    download_folder="downloads",
    max_workers=5,
    timeout=60,
    retry_count=3
)

# Загружаем из Excel
stats = downloader.download_from_excel(
    excel_file="image_links.xlsx",
    url_column="image_url",
    filename_column="query"
)

print(f"✅ Успешно: {stats['success']}")
print(f"❌ Ошибок: {stats['failed']}")
📊 Форматы входных файлов
Excel (.xlsx)
excel
| Query              | URL                                    |
|--------------------|----------------------------------------|
| Красивый закат     | https://example.com/sunset.jpg         |
| Горный пейзаж      | https://example.com/mountain.png       |
CSV (.csv)
csv
Query,URL
Красивый закат,https://example.com/sunset.jpg
Горный пейзаж,https://example.com/mountain.png
Текстовый файл (.txt)
text
https://example.com/image1.jpg
https://example.com/image2.png
https://example.com/image3.jpg
🔧 Настройка
Оптимизация работы парсера
python
parser = YandexImagesLinkParser(
    output_file="image_links.xlsx",
    file_format='excel'
)

# Настройка через параметры
parser.process_list(
    items=queries,
    batch_size=30  # Количество запросов до перезапуска драйвера
)
Оптимизация загрузчика
python
downloader = BatchFileDownloader(
    download_folder="downloads",
    max_workers=5,      # Количество потоков (1-10)
    timeout=60,         # Таймаут в секундах
    retry_count=3       # Количество попыток
)
🐛 Устранение проблем
Частые проблемы и решения
<details> <summary><b>❌ Ошибка: "ChromeDriver not found"</b></summary>
Решение: Установите ChromeDriver автоматически:

bash
pip install webdriver-manager
Или скачайте вручную с официального сайта.

</details><details> <summary><b>❌ Ошибка: "no module named 'distutils'"</b></summary>
Решение: Установите setuptools:

bash
pip install setuptools
</details><details> <summary><b>❌ Яндекс требует авторизацию</b></summary>
Решение:

Установите undetected-chromedriver:

bash
pip install undetected-chromedriver
Увеличьте задержки между запросами:

python
parser.random_delay(5, 10)  # Увеличенные задержки
Используйте разные User-Agent (встроено автоматически)

</details><details> <summary><b>❌ Медленная работа</b></summary>
Решение:

Уменьшите количество потоков загрузки

Увеличьте интервалы между запросами

Используйте batch_size для периодического перезапуска драйвера

</details>
📈 Советы по эффективному использованию
Парсинг ссылок:

Используйте файлы с запросами для массовой обработки

Устанавливайте batch_size=20-30 для стабильной работы

Проверяйте progress.txt для отслеживания прогресса

Загрузка изображений:

Начинайте с max_workers=3 и увеличивайте постепенно

Используйте --limit для тестирования перед массовой загрузкой

Проверяйте failed_downloads.json для повторной загрузки неудачных файлов

Обход ограничений:

Добавляйте случайные задержки между запросами

Используйте разные User-Agent

Обновляйте браузер и ChromeDriver

📝 Лицензия
MIT License - свободное использование, модификация и распространение.

🤝 Вклад в проект
Приветствуются:

🐛 Сообщения об ошибках

💡 Предложения по улучшению

🔧 Pull Requests

📧 Контакты
GitHub: [Ваш профиль]

Email: [Ваш email]

Issues: [Ссылка на Issues]

⚠️ Отказ от ответственности
Данный проект создан в образовательных целях. Автор не несет ответственности за:

Нарушение авторских прав

Нарушение условий использования сервисов

Неправомерное использование собранных данных

Используйте инструмент ответственно и соблюдайте законодательство вашей страны.
