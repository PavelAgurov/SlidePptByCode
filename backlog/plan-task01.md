# План реализации: LLM-агент для генерации презентаций

## Контекст

Нужно создать LLM-агента, который по текстовому описанию генерирует Python-код для создания PowerPoint-презентаций через python-pptx. Агент подключается к GPT-4.1 через OpenRouter, генерирует код, выполняет его, и при ошибках — автоматически исправляет.

**Вход:** текстовый файл с описанием презентации (например `data/data01.md`)  
**Выход:** файл `.pptx` в папке `.output/`, сгенерированный код в `.generated/`

---

## Структура проекта

```
SlidePptCode/
├── src/
│   ├── __init__.py
│   ├── main.py              # CLI точка входа + оркестрация пайплайна
│   ├── config.py             # Настройки через pydantic-settings
│   ├── models.py             # Pydantic-модели для structured output
│   ├── llm_client.py         # Обёртка над OpenAI клиентом для OpenRouter
│   ├── code_generator.py     # Генерация и исправление кода через LLM
│   ├── code_executor.py      # Выполнение кода в subprocess
│   ├── validator.py          # Валидация .pptx (кол-во слайдов, заголовки)
│   └── prompts.py            # Системные промпты и шаблоны
├── tests/
│   ├── __init__.py
│   ├── test_config.py
│   ├── test_models.py
│   ├── test_code_executor.py
│   ├── test_code_generator.py
│   ├── test_validator.py
│   ├── test_integration.py   # E2E тест с реальным LLM (маркер @integration)
│   └── fixtures/
│       ├── sample_task.md
│       └── sample_code.py
├── data/
│   └── data01.md
├── .generated/               # создаётся автоматически
├── .output/                  # создаётся автоматически
├── .logs/                    # создаётся автоматически
├── .env
├── .gitignore
├── requirements.txt
└── backlog/
    ├── task01.txt
    └── plan-task01.md
```

---

## Модули и ответственности

### 1. `src/config.py` — Конфигурация

```python
class Settings(BaseSettings):
    openrouter_api_key: str
    model_id: str = "gpt-4.1-mini"      # настраиваемый параметр
    base_url: str = "https://openrouter.ai/api/v1"
    temperature: float = 0
    max_retries: int = 3
    execution_timeout: int = 60
    generated_dir: Path = Path(".generated")
    output_dir: Path = Path(".output")
```

Использует `pydantic-settings` + `python-dotenv` для загрузки `.env`.

### 2. `src/models.py` — Pydantic-модели

- `GeneratedCode` — structured output от LLM: `code`, `explanation`, `output_filename`
- `CodeExecutionResult` — результат выполнения: `success`, `output_file`, `error_message`, `stdout`, `stderr`

### 3. `src/llm_client.py` — Клиент OpenRouter

Обёртка над `OpenAI(base_url=..., api_key=...)`. Метод `generate_structured()` — вызов `client.beta.chat.completions.parse()` с response_format=Pydantic модель.

### 4. `src/prompts.py` — Промпты

**Системный промпт** включает:
- Роль (эксперт python-pptx)
- Требования: код самодостаточный, сохранять в `.output/`, создавать директорию
- Паттерны python-pptx: добавление слайдов, заголовков, bullet points, форматирование
- Ограничения: только Python-код, без интерактивного ввода

**Промпт исправления ошибки** включает: оригинальный код, traceback, stderr/stdout.

### 5. `src/code_generator.py` — Генерация кода

Класс `CodeGenerator`:
- `generate_initial_code(task: str) -> GeneratedCode` — первая генерация
- `fix_code(code: str, error: CodeExecutionResult) -> GeneratedCode` — исправление
- Хранит историю сообщений для контекста ошибок

### 6. `src/code_executor.py` — Выполнение кода

Класс `CodeExecutor`:
- `save_code(code, filename) -> Path` — сохранение в `.generated/`
- `execute(code_file) -> CodeExecutionResult` — запуск через `subprocess.run()` с таймаутом
- Использует Python из `.venv/Scripts/python.exe`

### 7. `src/validator.py` — Валидация результата

Функция `validate_presentation(pptx_path, task_content) -> ValidationResult`:
- Проверяет что файл существует и не пустой
- Открывает .pptx через python-pptx
- Считает количество слайдов
- Проверяет наличие заголовков на слайдах
- Сравнивает кол-во слайдов с ожидаемым (парсит `## Slide N` из задания)

### 8. `src/main.py` — CLI + оркестрация

```
python src/main.py data/data01.md [--max-retries 3] [--model gpt-4.1] [--verbose]
```

Пайплайн:
1. Загрузить конфиг → создать директории
2. Прочитать файл задания
3. Сгенерировать код через LLM
4. Сохранить код → выполнить в subprocess
5. При ошибке — отправить ошибку в LLM → получить исправленный код → повторить (до max_retries)
6. При успехе — валидировать .pptx (кол-во слайдов)
7. Если валидация не прошла — попросить LLM исправить (это тоже считается ретраем)
8. Вывести результат

---

## Порядок реализации

### Шаг 1: Инфраструктура
- `src/__init__.py`, `src/config.py`, `src/models.py`
- Создание директорий `.generated/`, `.output/`

### Шаг 2: LLM-клиент и промпты
- `src/llm_client.py`, `src/prompts.py`

### Шаг 3: Генерация и выполнение
- `src/code_generator.py`, `src/code_executor.py`

### Шаг 4: Валидация
- `src/validator.py`

### Шаг 5: Оркестрация и CLI
- `src/main.py` — argparse + пайплайн с retry loop

### Шаг 6: Тесты
- Unit-тесты для каждого модуля
- Integration-тест с реальным LLM

---

## Тесты (pytest)

### Unit-тесты

| Файл | Что тестирует |
|---|---|
| `test_config.py` | Загрузка настроек, дефолтные значения, валидация |
| `test_models.py` | Валидация Pydantic-моделей, сериализация |
| `test_code_executor.py` | Сохранение кода, выполнение (mock subprocess), таймауты, обнаружение .pptx |
| `test_code_generator.py` | Построение сообщений, retry-логика (mock LLM) |
| `test_validator.py` | Валидация .pptx: подсчёт слайдов, парсинг задания |

### Integration-тест

`test_integration.py` (`@pytest.mark.integration`):
- Полный E2E пайплайн с реальным LLM
- Простое задание на 2-3 слайда
- Проверка: файл .pptx создан, можно открыть, содержит слайды

### Фикстуры

- `tests/fixtures/sample_task.md` — простое задание для тестов
- `tests/fixtures/sample_code.py` — рабочий python-pptx код для тестирования executor без LLM

### Запуск

```bash
# unit-тесты
python -m pytest tests/ -v --ignore=tests/test_integration.py

# integration
python -m pytest tests/test_integration.py -v -m integration

# всё
python -m pytest tests/ -v
```

---

## Верификация

1. **Unit-тесты**: `python -m pytest tests/ -v --ignore=tests/test_integration.py` — все зелёные
2. **Ручной прогон**: `python src/main.py data/data01.md --verbose` — создаётся `.output/*.pptx`
3. **Открыть .pptx** — проверить что 8 слайдов, есть заголовки и контент
4. **Тест ретрая**: убедиться что при ошибке LLM получает traceback и исправляет
5. **Integration-тест**: `python -m pytest tests/test_integration.py -v` — проходит
