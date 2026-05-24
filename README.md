# Yandex Wordstat API Agent
Автоматизированный инструмент для сбора и анализа поискового спроса через Yandex Search API v2.
## Возможности
- Сбор статистики поисковых запросов через официальный Yandex Search API v2
- Анализ динамики спроса по месяцам
- Генерация Excel-отчёта с листами: Данные, Сводка, График, Голубые океаны, Растущие рынки
- Настройка временного диапазона анализа
- Поддержка множественных запросов
## Требования
- Python 3.8+
- API-ключ Яндекс.Облака
- Зависимости из requirements.txt
## Установка
```bash
git clone https://github.com/Dimitriy1204/ваш-репозиторий
cd ваш-репозиторий
pip install -r requirements.txt
```
## Настройка
Скопируйте `.env.example` в `.env` и заполните своими данными:
```
YANDEX_API_KEY=ваш_ключ
FOLDER_ID=ваш_folder_id
SEARCH_QUERIES=купить оборудование,купить станок
DATE_FROM=2022-01
DATE_TO=2024-12
```
## Запуск
```bash
python analyzer.py
```
Результат сохраняется в `yandex_analysis.xlsx`.
## Структура проекта
- `analyzer.py` — основной скрипт
- `test_api.py` — тест API
- `check_excel.py` — проверка отчёта
- `requirements.txt` — зависимости
## Контакты
Автор: Dimitriy1204
GitHub: https://github.com/Dimitriy1204
