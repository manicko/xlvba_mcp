---
id: ru-xlvba-development-and-maintenance
domain: ru-11-development-and-maintenance
tags:
  - development
  - maintenance
  - architecture
  - testing
  - contributing
related:
  - ru-xlvba-overview
  - ru-xlvba-installation-and-setup
  - ru-xlvba-configuration
  - ru-xlvba-architecture
  - ru-xlvba-mcp-tools-reference
  - ru-xlvba-excel-capabilities
  - ru-xlvba-workflows-and-examples
  - ru-xlvba-security
  - ru-xlvba-limitations-and-compatibility
  - ru-xlvba-troubleshooting
---

## Назначение

Документ предназначен для **сопровождающих и контрибьюторов** `xlvba`. Он объясняет структуру
репозитория, границы слоёв, как добавить новый MCP-инструмент, не нарушив архитектуру, контракт
обработки ошибок, а также процесс тестирования, логирования и релиза.

Если вы хотите только использовать сервер, начните с [Обзор](01-overview.md) и
[Установка и настройка](02-installation-and-setup.md).

Управление документацией (контракт frontmatter, kebab-case, относительные ссылки, пороги разделения)
определено в [doc-maintenance-rules](../00-overview/doc-maintenance-rules.md); следуйте ему для каждого
изменения в `docs/`.

## Структура репозитория

| Путь | Ответственность |
|------|-----------------|
| `pyproject.toml` | Метаданные пакета, зависимости, консольный скрипт `excel-vba-mcp`, конфигурация инструментов (ruff, mypy, pytest) |
| `uv.lock` | Зафиксированные версии зависимостей |
| `README.md` | Точка входа, требования, установка/запуск, конфигурация, каталог инструментов |
| `src/excel_vba_mcp/__main__.py` | Консольная точка входа (`mcp.run(transport="stdio")`) |
| `src/excel_vba_mcp/server.py` | `MCPServer`, lifespan, регистрация инструментов/ресурсов, граница структурированных ошибок, логирование |
| `src/excel_vba_mcp/config.py` | Типизированная конфигурация окружения (Pydantic) |
| `src/excel_vba_mcp/api/` | Тонкие асинхронные MCP-обработчики инструментов (валидация + делегирование) |
| `src/excel_vba_mcp/application/` | Сессия, снапшот, модель книги, листы, таблицы, диапазоны, сводные, имена, фигуры, элементы, VBA, компиляция, макрос, тест-раннер |
| `src/excel_vba_mcp/adapters/excel/` | Однопроточный STA COM-воркер и адаптеры Excel/VBIDE |
| `src/excel_vba_mcp/adapters/windows_ui/` | Read-only адаптер UI компиляции VBE (`comtypes`) |
| `src/excel_vba_mcp/domain/` | Модели, статусы/enum'ы, стабильные коды ошибок, декларативная тест-схема |
| `src/excel_vba_mcp/security/` | Approval, аудит, path policy, macro policy |
| `src/excel_vba_mcp/storage/` | Хранилище снапшотов, хранилище артефактов, репозиторий прогонов |
| `tests/unit/` | Без COM / Excel |
| `tests/integration_excel/` | Живые тесты Excel COM (маркер `integration_excel`) |
| `tests/fixtures/` | Фикстуры книг и детерминированные генераторы |
| `docs/` | Документация проекта |
| `.ai/` | Планы, аудиты, контекст, логи выполнения |
| `.kilo/` | Агенты, команды, правила |

## Основные модули исходников

| Модуль | Ответственность |
|--------|-----------------|
| `server.py` | Собирает `MCPServer`, владеет рантаймом через `lifespan`, регистрирует 88 инструментов и 4 ресурса, повторно поднимает структурированные `ToolError`, настраивает логирование в `stderr` |
| `config.py` | Загружает `EXCEL_VBA_MCP_LOG_LEVEL` и `EXCEL_VBA_MCP_ACCESS_MODE` в Pydantic `Config`; валидирует их совместно |
| `api/*_tools.py` | Один модуль на группу инструментов; каждый обработчик тонкий: валидирует аргументы, делегирует сервису, возвращает DTO |
| `application/session_service.py` | Открытие/закрытие/сохранение рабочих копий, блокировка на путь, статус сессии, восстановление |
| `application/snapshot_service.py` | Иммутабельные снапшоты и восстановление |
| `application/*_service.py` | Листы, таблицы, диапазоны, сводные, имена, фигуры, элементы |
| `application/vba_project_service.py` | Чтение проекта/компонентов/исходников VBA, экспорт, защищённые замена/вставка |
| `application/compile_service.py` | Fail-closed диагностика компиляции |
| `application/macro_service.py` | Инвентарь макросов, реестр `background_safe`, preflight R13, диспетчеризация `Application.Run` |
| `application/test_runner.py` | Декларативный тест-раннер над одноразовыми копиями |
| `application/workbook_model_service.py` | Агрегированная модель книги |
| `adapters/excel/com_worker.py` | Единственный STA-поток, FIFO-очередь, мягкие таймауты, read-only-повтор, восстановление |
| `adapters/excel/app_adapter.py` | Скрытый `Excel.Application`, `DispatchEx`, узко ограниченные настройки, завершение по pid |
| `adapters/excel/*.py` | Адаптеры листов/таблиц/диапазонов/сводных/VBIDE/фигур/элементов/макросов |
| `adapters/windows_ui/vbe_compile_ui.py` | Read-only обнаружение команд VBE и наблюдение диалогов (`comtypes`) |
| `domain/models.py` | Pydantic DTO и пределы размера |
| `domain/statuses.py` | Закрытые словaries статусов и enum'ы |
| `domain/errors.py` | Иерархия `ExcelVbaMcpError`, `ErrorCode`, сопоставление HRESULT |
| `domain/test_spec.py` | Декларативная тест-схема (закрытые enum'ы действий/проверок) |
| `security/path_policy.py` | Allowlist и канонизация |
| `security/approval.py` | Политика approval `full` / `restricted` |
| `security/audit.py` | Append-only JSONL-аудит с редактированием |
| `security/macro_policy.py` | Режимы `AutomationSecurity` и диагностика макросов |
| `storage/snapshot_store.py` | Иммутабельное хранилище снапшотов |
| `storage/artifact_store.py` | Артефакты и доказательства |
| `storage/run_repository.py` | Сохранённые декларативные прогоны |

## Среда разработки

```powershell
uv sync --locked
```

- `.python-version` фиксирует **3.14**; нижний предел проекта — **3.11**
  (`requires-python = ">=3.11,<3.15"`).
- Группа `dev` (`pytest`, `ruff`, `mypy`) устанавливается по умолчанию.
- Запуск сервера:

```powershell
uv run excel-vba-mcp
```

- Консольный скрипт также доступен по пути `<repo>\.venv\Scripts\excel-vba-mcp.exe` после `uv sync`.

## Как добавить новый MCP-инструмент

Следуйте существующим слоям. Инструмент — это не одна функция; это DTO + метод адаптера + метод
сервиса + тонкий обработчик + регистрация + тесты + документация.

1. **Добавьте DTO и enum'ы в `domain/`.** Определите модели запроса/ответа в `domain/models.py` и
   любой новый словарь статусов в `domain/statuses.py`. Доменные модели — это Pydantic `BaseModel`
   и должны оставаться без COM.
2. **Добавьте метод адаптера в `adapters/excel/`.** COM остаётся на STA-воркере. Метод адаптера
   выполняется внутри отправленного callable и возвращает **только DTO** (никогда COM-прокси).
3. **Добавьте метод сервиса в `application/`.** Сервис должен проходить через `worker.submit(...)` и
   обеспечивать цепочку защиты: **writable-сессия → снапшот → approval → проверки, специфичные для
   операции → перечитывание postcondition + аудит**. Сервис никогда не держит COM-прокси.
4. **Добавьте тонкий асинхронный обработчик в `api/<group>_tools.py`.** Дайте ему Pydantic-модель
   входа с `extra="forbid"`, валидируйте аргументы, делегируйте сервису и верните DTO. Слой API не
   должен импортировать `pythoncom`/`win32com` и не должен трогать COM.
5. **Зарегистрируйте его в `server.py` `_register_tools`.** Используйте
   `server.add_tool(handler, name="...", description=_DESC_...)` со стабильным именем и константой
   описания `_DESC_*`. Инструмент без зависимостей (например, `server_health`) может использовать
   `@mcp.tool()` напрямую.
6. **Добавьте тесты.** Модульные тесты идут в `tests/unit` (без COM/Excel). Если инструмент трогает
   Excel, добавьте интеграционный тест в `tests/integration_excel` (маркер `integration_excel`).
7. **Обновите документацию.** Добавьте инструмент в
   [Справочник инструментов MCP](05-mcp-tools-reference.md) и, если он добавляет возможность Excel,
   в [Возможности Excel](06-excel-capabilities.md).

## Соглашения по реализации

- **Только английский.** Все комментарии, логи, docstrings, сообщения об ошибках и документация — на
  английском.
- **Единый источник истины на факт.** Факт живёт ровно в одном модуле или документе.
- **Границы слоёв.**
  - `api` не должен импортировать `pythoncom`/`win32com` и не должен трогать COM.
  - Сервисы никогда не держат COM-прокси; каждая операция — это `await worker.submit(...)`.
  - `domain` не зависит от COM; HRESULT — только целое число.
- **Граница DTO.** `_ensure_dto` валидирует каждое возвращаемое воркером значение (Pydantic-модель,
  примитив или контейнер из них). COM-прокси или любой другой несериализуемый объект поднимает
  исключение и фиксируется как `PROXY_LEAK_SUSPECTED`.
- **Стабильный словарь `ErrorCode`.** Коды только добавляются; никогда не удаляйте и не меняйте
  существующий код.
- **Pydantic `extra="forbid"`** на входах инструментов.
- **Явный `shift`, никогда не угадывается.** `range_insert` / `range_delete` отклоняют отсутствующий
  или неизвестный shift.
- **Нет слепого повтора мутаций.** Повторяются только вызовы `READ_ONLY`.
- **Аудит каждой операции.** Каждый мутирующий и исполняющий путь записывает аудит-запись со своим
  `operation_class` и результатом.
- **Стабильные имена MCP-инструментов.** Имена инструментов — публичный контракт; не переименовывайте
  их бездумно.

## Требования к обработке ошибок

- Поднимайте подклассы `ExcelVbaMcpError` со стабильным `ErrorCode` (`domain/errors.py`).
- Включайте `hresult`, когда ошибка пришла из COM; `hresult_hex` выводится.
- Предоставляйте структурированный словарь `context` (путь, `workbook_id`, `reason`, `section`,
  `token`, `class`, …).
- Никогда не протекалайте COM-объекты через границу; воркер пересекают только DTO.
- Сбои валидации аргументов сопоставляются с `INVALID_RANGE` и `reason=invalid_arguments`
  (см. `server._validation_error_as_domain`).
- **Никогда не заявляйте об успехе при `indeterminate`-компиляции.** `vba_compile` возвращает
  `status="indeterminate"` с `diagnostics.is_compile=false` на валидированной сборке.

## Подход к тестированию

- **Модульные тесты** (`tests/unit`) выполняются без COM/Excel и должны оставаться герметичными.
- **Интеграционные тесты** (`tests/integration_excel`) управляют живым экземпляром Excel COM. Они
  помечены `integration_excel` и пропускаются, когда Excel недоступен. Они выполняются последовательно,
  поскольку экземпляр Excel один.
- **Фикстуры** строятся `tests/fixtures/build_fixtures.py` и `tests/fixtures/build_phase2_fixtures.py`.
  Пересобирайте их, когда бинарные фикстуры отсутствуют или устарели. `LockedProject.xlsm` и
  `TableNamesProject.xlsm` не закоммичены, поэтому эти тесты пропускаются с явной причиной.

```powershell
uv run ruff check src tests
uv run ruff format --check src tests
uv run mypy src
uv run pytest tests/unit
uv run pytest -m integration_excel
```

Запустите все четыре статические/модульные команды перед предложением изменения; запускайте
интеграционный набор на Windows с Excel, когда изменение затрагивает поведение COM.

## Логирование и отладка

- Логи идут **только в `stderr`** на фиксированном уровне `INFO` (`server._configure_logging()`);
  `stdout` зарезервирован для протокольного трафика MCP.
- `EXCEL_VBA_MCP_LOG_LEVEL` разбирается и валидируется, но **не применяется** в этой ревизии; не
  полагайтесь на него для изменения многословности.
- Аудит-след (`audit.jsonl`) — авторитетная запись того, что произошло.
- Используйте `server_health` (liveness) и `excel_health` (состояние воркера/сессии) для диагностики
  без обращения к COM.
- Структурированные ошибки несут `code`/`context`; предпочитайте их разбору текста логов.

## Рабочий процесс Git и релиза

- Осмотрите историю на предмет стиля сообщений коммитов:

```powershell
git log --oneline -10
```

- Используйте короткие, конвенциональные темы (например, `docs: ...`, `fix: ...`, `feat: ...`,
  `test: ...`).
- Стейджьте и коммитьте **только нужные файлы**; сначала просмотрите `git status` и `git diff`.
  Никогда не коммитьте секреты.
- Не амендьте и не ребейзьте общую историю. Если коммит не проходит хук, исправьте проблему и создайте
  новый коммит, а не амендьте проваленный.
- Не коммитьте рантайм-данные (`<temp>/excel_vba_mcp/`) или артефакты сборки.

## Поддержка зависимостей

- Отредактируйте `pyproject.toml`, затем осознанно обновите lock и окружение:

```powershell
uv lock
uv sync
```

- Рантайм-зависимости: `mcp`, `pydantic`, `pywin32`, `comtypes`. Зависимости разработки: `pytest`,
  `ruff`, `mypy`.
- `pyproject.toml` объявляет `mcp>=2.3,<3`; lock разрешает 2.3.0. Код использует API `MCPServer`
  SDK v2 (`mcp.server.mcpserver`).
- Фиксируйте точные версии через `uv.lock` и ревьюьте обновления зависимостей; обновление
  `pywin32`/`comtypes` может изменить поведение COM на валидированной сборке.

## Связанные документы

- [Обзор](01-overview.md) — что такое сервер.
- [Установка и настройка](02-installation-and-setup.md) — настройка окружения.
- [Конфигурация](03-configuration.md) — переменные окружения.
- [Архитектура](04-architecture.md) — слои, воркер, цепочка защиты.
- [Справочник инструментов MCP](05-mcp-tools-reference.md) — контракт инструментов.
- [Возможности Excel](06-excel-capabilities.md) — матрица возможностей.
- [Рабочие процессы и примеры](07-workflows-and-examples.md) — рецепты использования.
- [Безопасность](08-security.md) — модель безопасности.
- [Ограничения и совместимость](09-limitations-and-compatibility.md) — известные пределы.
- [Устранение неполадок](10-troubleshooting.md) — диагностика сбоев.
- [Правила документации](../00-overview/doc-maintenance-rules.md) — управление документацией.