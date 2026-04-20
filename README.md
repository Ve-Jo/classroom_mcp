# Google Classroom MCP Server — Enhanced Fork

> MCP-сервер для подключения Claude, Codex и других AI-ассистентов к Google Classroom.  
> Форк проекта [SalShah20/classroom_mcp](https://github.com/SalShah20/classroom_mcp) с глубокой переработкой: OCR на базе PaddleOCR, персистентный кэш материалов, улучшенный поиск и безопасное хранение авторизации.

---

## Что это

Сервер работает по протоколу [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) поверх stdio и предоставляет AI-ассистенту доступ к данным Google Classroom в режиме только чтения: курсы, задания, объявления, оценки, вложения и полный текстовый контент материалов.

AI-ассистент может:
- Искать задания, курсы и объявления с релевантным ранжированием
- Извлекать текст из документов, PDF, презентаций, таблиц, изображений и Google Forms
- Распознавать текст на сканированных PDF и фото с помощью OCR (в т.ч. кириллица)
- Просматривать дашборд с предстоящими и просроченными заданиями

---

## Отличия от оригинальной версии

| Возможность | Оригинал (SalShah20) | Этот форк |
|---|---|---|
| OCR для сканированных PDF | Нет | PaddleOCR v5 server models с рендером страниц в 3x |
| OCR для изображений | Нет | PaddleOCR + OpenCV-препроцессинг (grayscale + adaptive threshold) |
| OCR для встроенных изображений в DOCX/PPTX/Google Docs | Нет | Автоматическое извлечение изображений из ZIP-архивов и HTML-экспортов |
| Распознавание кириллицы | Нет | Автодетект кириллицы с переключением на кириллическую модель |
| GPU-ускорение OCR | Нет | DirectML (GPU) через ONNX Runtime с fallback на CPU |
| Персистентный кэш материалов | Нет (только in-memory) | SQLite (WAL-mode) с TTL 6 ч, поиск по содержимому, пагинация |
| Извлечение контента Google Forms | Нет | Forms API + fallback на парсинг публичной страницы + OCR изображений |
| Безопасное хранение токенов | tokens.json (plaintext) | DPAPI-шифрование (Windows), файл auth.secure.json |
| Унифицированный поиск | Разрозненные tools | Единый search-движок с ranking, facets, cursor-пагинацией, автодополнением |

---

## Инструменты

### Основные

| Инструмент | Описание |
|---|---|
| `search` | Унифицированный поиск по заданиям, курсам и объявлениям с ранжированием, фасетами и курсорной пагинацией |
| `suggest_search_terms` | Автодополнение и расширение запросов для AI-потоков |
| `get_assignment_material_text` | Извлечение текста из вложений заданий: Google Docs/Sheets/Slides, PDF, DOCX, PPTX, XLSX, изображения (OCR), Google Forms |
| `search_material_cache` | Поиск по кэшированному тексту извлечённых материалов |
| `read_material_cache` | Пагинационное чтение кэшированного текста по `docRef` |
| `list_material_cache` | Список кэшированных документов с метаданными |
| `clear_material_cache` | Очистка кэша материалов |
| `get_dashboard` | Компактный обзор: активные курсы, предстоящие/просроченные задания, сводка оценок |

### Legacy (обратная совместимость)

`courses`, `course-details`, `assignments`, `list_courses`, `get_course`, `list_coursework`, `get_coursework`, `list_submissions`, `list_announcements`, `get_upcoming_assignments`, `get_missing_assignments`, `get_assignments`, `calculate_grade`, `get_assignment_materials`, `get_grades`, `search_assignments`

---

## Установка и настройка

### Требования

- Node.js v16+
- Google Cloud проект с включённым Classroom API
- (Опционально) GPU с поддержкой DirectML для ускорения OCR на Windows

### Шаг 1 — Клонирование и установка

```bash
git clone https://github.com/Ve-Jo/classroom_mcp.git
cd classroom_mcp
npm install
npm run build
```

### Шаг 2 — Включить Google Classroom API

1. Открой [console.cloud.google.com](https://console.cloud.google.com/)
2. Создай/выбери проект
3. **APIs & Services > Library** → включи **Google Classroom API**
4. Для извлечения Google Forms — включи **Google Forms API** (опционально)

### Шаг 3 — Создать OAuth-credentials

1. **APIs & Services > Credentials** → **Create Credentials > OAuth 2.0 Client ID**
2. Настрой consent screen (External + свой тестовый пользователь)
3. Тип: **Desktop app**
4. Скачай JSON → сохрани как `credentials.json` в корне проекта

### Шаг 4 — Авторизация

```bash
npm run setup-auth
```

Процесс:
1. Читает `credentials.json`
2. Запускает callback-сервер на `localhost`
3. Выводит URL для авторизации
4. Автоматически захватывает код (или принимает вставленный URL)

Результат:
- Шифрованные данные → `auth.secure.json` (Windows DPAPI)
- Настройки → `.env` + `GOOGLE_AUTH_STORE`
- Если шифрование недоступно — fallback на plaintext-режим через `.env`

### Шаг 5 — Подключение к Claude Desktop

Конфиг:
- Windows: `%APPDATA%\Claude\claude_desktop_config.json`
- macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "google-classroom": {
      "command": "node",
      "args": ["C:/path/to/classroom_mcp/dist/index.js"],
      "env": {
        "GOOGLE_AUTH_STORE": "C:/path/to/classroom_mcp/auth.secure.json",
        "GOOGLE_MATERIAL_CACHE_DB": "C:/path/to/classroom_mcp/material-cache.sqlite"
      }
    }
  }
}
```

`GOOGLE_AUTH_STORE` — путь к файлу, созданному `setup-auth`.  
`GOOGLE_MATERIAL_CACHE_DB` — путь к SQLite-базе кэша материалов.

Legacy-режим через переменные окружения:
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN`

Перезапусти Claude Desktop после сохранения конфига.

---

## OCR-движок

Встроенный OCR на базе **PaddleOCR v5** с ONNX Runtime:

- **Детекция текста**: PP-OCRv5 server detection model
- **Распознавание**: PP-OCRv5 server recognition model (английский) + мобильная модель для кириллицы
- **Препроцессинг**: OpenCV (grayscale + adaptive threshold)
- **GPU-ускорение**: DirectML через ONNX Runtime, fallback на CPU
- **Автодетект языка**: если >15% кириллических символов — автоматически переключается на кириллическую модель
- **Таймаут**: 2 минуты на одно изображение
- **Лимиты**: до 10 изображений на документ, до 10 МБ на изображение

Модели загружаются автоматически при первом запуске (если нет в `models/`). Для офлайн-использования — помести ONNX-файлы в папку `models/`.

### Поддерживаемые форматы извлечения

| Формат | Метод | OCR |
|---|---|---|
| Google Docs | Export → text/plain + HTML для изображений | Да (embedded images) |
| Google Sheets | Export → CSV | Нет |
| Google Slides | Export → text/plain | Нет |
| PDF | pdf-parse + рендер страниц через pdfjs-dist | Да (сканированные PDF) |
| DOCX | mammoth + ZIP-извлечение изображений | Да (embedded images) |
| PPTX | Парсинг XML-слайдов + ZIP-извлечение изображений | Да (slide images) |
| XLSX | SheetJS → CSV | Нет |
| Изображения (PNG, JPEG, ...) | Напрямую | Да |
| Google Forms | Forms API или парсинг публичной страницы | Да (form images) |

---

## Кэш материалов

Извлечённый текст кэшируется в **SQLite** (WAL-mode) с TTL 6 часов:

- Автоматическая очистка истёкших записей
- Полнотекстовый поиск по кэшу (`search_material_cache`)
- Пагинационное чтение через `docRef` (`read_material_cache`)
- Принудительная очистка через `clear_material_cache`

---

## Справочник API

### Unified Search

`search(query?, entityTypes?, courseIds?, states?, dueFrom?, dueTo?, missingOnly?, gradedOnly?, limit?, cursor?, sort?, forceRefresh?)`

- `entityTypes`: `assignments | courses | announcements`
- `sort`: `relevance | dueDate | updatedAt`
- `cursor`: opaque-токен из предыдущего ответа
- `forceRefresh`: принудительное обновление индекса

Ранжирование учитывает:
- Совпадение в заголовке (+0.55) и содержимом (+0.25)
- Token overlap (+до 0.15)
- Близость дедлайна (+до 0.14)
- Статус missing (+0.2)
- Свежесть обновления (+0.1)

### Примеры запросов

1. Просроченные задания в курсе на 14 дней вперёд  
   `entityTypes=['assignments']`, `courseIds=['<id>']`, `missingOnly=true`, `dueTo=<+14d>`, `sort='dueDate'`

2. Последние объявления  
   `entityTypes=['announcements']`, `sort='updatedAt'`

3. Поиск по ключевому слову  
   `entityTypes=['assignments']`, `query='project report'`, `sort='relevance'`

4. Только оценённые задания  
   `entityTypes=['assignments']`, `gradedOnly=true`

5. Пагинация  
   Вызови с `limit`, затем продолжи через `nextCursor`

---

## Google API Scopes

- `classroom.courses.readonly`
- `classroom.course-work.readonly`
- `classroom.student-submissions.me.readonly`
- `classroom.announcements.readonly`
- `classroom.courseworkmaterials.readonly`
- `classroom.topics.readonly`
- `drive.readonly` (для `get_assignment_material_text`)
- `forms.body.readonly` (для извлечения Google Forms)

Все scope — только чтение. Запись не запрашивается.

---

## Структура проекта

```
classroom_mcp/
├── src/
│   ├── index.ts          # MCP-сервер, поиск, OCR, кэш
│   ├── setup-auth.ts     # Интерактивная OAuth-авторизация
│   └── auth-store.ts     # DPAPI-шифрование токенов
├── models/               # ONNX-модели PaddleOCR (автозагрузка)
├── dist/                 # Скомпилированный JS
├── credentials.json      # OAuth credentials (не в git)
├── auth.secure.json      # Шифрованные токены (не в git)
├── material-cache.sqlite # Кэш материалов (не в git)
├── .env                  # Настройки окружения (не в git)
├── inject-credentials.js # Инъекция кредов в dist при npm publish
├── package.json
├── tsconfig.json
└── README.md
```

---

## Решение проблем

**`credentials.json not found`**  
Положи файл в корень проекта.

**`No refresh token received`**  
Отзови доступ на [myaccount.google.com/permissions](https://myaccount.google.com/permissions), затем перезапусти `npm run setup-auth`.

**`Google Classroom API not initialized`**  
Установи `GOOGLE_AUTH_STORE` на абсолютный путь `auth.secure.json`.

**OCR не работает / медленно**  
- Убедись, что в `models/` есть ONNX-файлы (или есть интернет для автозагрузки)
- Для GPU-ускорения нужна поддержка DirectML (Windows, совместимый GPU)
- Без GPU — используется CPU, что медленнее

**Сервер не появляется в Claude Desktop**  
- Проверь путь в `args` → `dist/index.js`
- Запусти `npm run build`
- Перезапусти Claude Desktop

---

## Безопасность

- Все Google API scope — только чтение
- Access-токены обновляются автоматически
- Рекомендуемый режим: refresh token в `auth.secure.json` зашифрован через Windows DPAPI
- Fallback-режим: refresh token в `.env` для совместимости
- Токены остаются локально на машине

---

## Лицензия

MIT — как и [оригинальный проект](https://github.com/SalShah20/classroom_mcp).
