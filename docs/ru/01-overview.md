---
id: ru-xlvba-overview
domain: ru-01-overview
tags:
  - overview
  - excel
  - vba
  - mcp
  - windows
related:
  - ru-xlvba-installation-and-setup
  - ru-xlvba-configuration
  - ru-xlvba-architecture
  - ru-xlvba-mcp-tools-reference
  - ru-xlvba-excel-capabilities
  - ru-xlvba-workflows-and-examples
  - ru-xlvba-security
  - ru-xlvba-limitations-and-compatibility
  - ru-xlvba-troubleshooting
  - ru-xlvba-development-and-maintenance
---

## Назначение

Документ описывает, что такое `xlvba` (пакет `excel-vba-mcp`), для кого он предназначен,
что он делает и чего он не делает. Это точка входа в документацию проекта; детали
установки, конфигурации и отдельных групп инструментов вынесены в дочерние документы.

## Что такое xlvba

`xlvba` — это локальный **MCP-сервер**, который работает по транспорту **STDIO** и
управляет реальным, **скрытым** экземпляром Microsoft Excel через COM на Windows.
Сервер не имеет собственного пользовательского интерфейса: он публикует набор MCP-инструментов
и ресурсов, которые вызывает MCP-клиент (например, Kilo или Claude Desktop).

Ключевые факты:

- Пакет: `excel-vba-mcp`, версия `0.1.0` (лицензия Proprietary).
- Сервер регистрирует **43 инструмента** и **5 read-only ресурсов** (четыре
  per-workbook URI-шаблона плюс статический `xlvba://glossary`). Плановый read-only инструмент
  `xlvba_environment` (режим доступа / политика / сборка / возможности) ещё не добавлен; enum-схемы
  запросов для `range_set_conditional_formatting`, `pivot_field_set_layout` и
  `pivot_field_apply_filter` будут сокращены (pruned) до поведения рантайма (плановые ссылки).
- Транспорт — MCP поверх STDIO: сообщения протокола идут в `stdout`, диагностика и
  трассировки — в `stderr`.
- Экземпляр Excel создаётся через `win32com.client.DispatchEx` и остаётся `Visible=False`
  на всём протяжении жизни процесса.
- Имена инструментов, которые публикует сервер, клиенты показывают с префиксом
  (например, в Kilo — `xlvba_excel_open`).

Подробное описание инструментов приведено в [Справочник инструментов MCP](05-mcp-tools-reference.md),
возможностей Excel — в [Возможности Excel](06-excel-capabilities.md).

## Ключевые возможности

- **Сессии книг и рабочие копии.** `excel_open` копирует исходный файл в каталог `work/`
  и открывает только копию; оригинал никогда не изменяется. По умолчанию сессия доступна для
  записи и выполнения (`mode="execute"`, `read_only=False`); передайте `mode="inspect"` и/или
  `read_only=True` для сессии с отключёнными макросами или только для чтения.
- **Read-only модель книги.** `workbook_get_model` агрегирует листы, таблицы, имена, фигуры,
  элементы управления и (опционально) VBA-компоненты и процедуры.
- **Листы.** `sheets_list`, `sheet_get`, а также мутирующие операции `sheet_add`,
  `sheet_rename`, `sheet_copy`, `sheet_delete`, `sheet_set_visibility`.
- **Таблицы (ListObject).** `tables_list`, `table_get`, `table_read_rows`,
  `table_append_rows`, `table_insert_row`, `table_update_rows`, `table_delete_rows`,
  `table_add_column`, `table_delete_column`, `table_create`, `table_resize`.
- **Диапазоны и ячейки.** Чтение и запись значений и формул, чтение и частичная запись
  форматирования, проверка данных (data validation), условное форматирование, а также
  `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge`.
- **PivotTables (только существующие).** Перечисление, чтение метаданных и кэшей, обновление,
  управление полями (ориентация, позиция, агрегация, свойства), фильтры через видимые элементы
  и очистка фильтров. Создание и удаление PivotTables/PivotCaches не поддерживается.
- **Имена (Named Ranges).** `names_list`, `name_get` и мутирующие `name_create`,
  `name_update`, `defined_name_delete`.
- **Фигуры и элементы управления.** `shapes_list`, `shape_get`, `shape_set_on_action`,
  `controls_list`, `control_get`, `control_set_on_action`.
- **VBA.** Инвентаризация проекта и компонентов, чтение исходников, поиск по коду, экспорт проекта,
  чтение ссылок, а также защищённые (guarded) замена и вставка исходников.
- **Диагностика компиляции (fail-closed).** `vba_compile_diagnostics` и `vba_get_compile_diagnostics`
  возвращают частичные диагностические данные, но никогда не сообщают ложный успех.
- **Макросы (под гейтом).** `macro_list`, `macro_run`, `macro_get_last_result`,
  `macro_classify_background`; запуск требует execute-режима, классификации `background_safe`
  для текущего хэша книги и ревизии исходников VBA и прохождения preflight интерактивного UI.
  В режиме `full` (по умолчанию) отдельный грант на запуск не нужен; в режиме `restricted`
  также действует approval-гейт.
- **Снапшоты.** `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff`
  с манифестом и SHA-256.
- **Декларативный тестовый раннер.** `test_suite` (действия `run` / `get-run`)
  с закрытым словарём действий и проверок. `test_suite` `run` — единственная каноническая точка входа
  (принимает mapping или typed-spec, плюс `background`).
- **Аудит и безопасность.** Журнал `audit.jsonl` (append-only), allowlist путей и
  политика доступа: полный доступ по умолчанию, режим `restricted` (deny-by-default approval)
  как явный opt-in.

Каждая группа описана подробнее в [Возможности Excel](06-excel-capabilities.md),
угрозы и границы — в [Архитектура](04-architecture.md) и [Возможности Excel](06-excel-capabilities.md).

## Как ИИ-агент взаимодействует с Excel через xlvba

Агент никогда не видит видимого окна Excel и не управляет GUI. Вместо этого работает цикл:

1. MCP-клиент отправляет вызов инструмента по STDIO (JSON-RPC).
2. Сервер валидирует аргументы (Pydantic-модели, пути через path policy, approval-гейт).
3. Запрос передаётся в единственный выделенный **STA-поток** (`ExcelComWorker`), где
   выполняется весь доступ к COM.
4. Код Excel выполняет операцию в скрытом, неактивируемом экземпляре Excel.
5. Результат возвращается как DTO (типизированные данные, без «сырых» COM-объектов),
   затем сериализуется в ответ MCP.

Таким образом весь доступ к COM сериализован на одном STA-потоке; инструменты никогда не
обращаются к COM напрямую, а между воркером и остальным кодом пересекаются только DTO.
Обработка ошибок и восстановление воркера описаны в [Архитектура](04-architecture.md).

## Типовые сценарии

1. **Осмотр книги.** `excel_open` (режим `inspect`, `read_only=true`) → `workbook_get_model`
   или `sheets_list` → `excel_close`.
2. **Чтение диапазона.** `excel_open` → `range_read_values` / `range_read_formulas` /
   `range_get_format` → `excel_close`.
3. **Инвентаризация VBA.** `excel_open` (inspect) → `vba_project_info` / `vba_list_components`
   / `vba_list_procedures` → `excel_close`.
4. **Снапшот перед изменением.** `snapshot_create` (для открытой книги используется
   `Workbook.SaveCopyAs`, для закрытой — копирование файла).
5. **Экспорт VBA.** `vba_export_project` с сохранением манифеста и хэшей файлов.
6. **Запуск декларативного теста.** `test_suite` (`action="run"`) → при необходимости
   `test_suite` (`action="get-run"`) для чтения сохранённого прогона.
7. **Диагностика состояния компиляции.** `vba_compile_diagnostics` / `vba_get_compile_diagnostics` —
   получение статуса и частичных диагностических данных без запуска UI.

Пошаговые примеры с аргументами приведены в [Рабочие процессы и примеры](07-workflows-and-examples.md).

## Чего xlvba НЕ делает

- Не автоматизирует GUI и не показывает видимого окна Excel.
- Не создаёт новые книги (`excel_open` открывает только существующие файлы).
- Не создаёт и не удаляет PivotTables/PivotCaches; не создаёт диаграммы, сводные диаграммы и слайсеры.
- Не предоставляет инструмент выдачи approval-грантов или редактирования allowlist через MCP.
  В режиме `full` (по умолчанию) грант не нужен (мутации и execute-режим авторизованы); в режиме
  `restricted` действует deny-by-default и мутирующие/исполняющие операции отклоняются с
  `APPROVAL_DENIED`.
- Не сообщает об успешной компиляции на валидированной сборке: `vba_compile_diagnostics` всегда возвращает
  `status=indeterminate` (`reason=NON_ACTIVATING_COMPILE_UNAVAILABLE`).
- Не выполняет макросы без классификации `background_safe`, привязанной к текущему хэшу книги и
  ревизии исходников VBA; классификацию задаёт `macro_classify_background`. Интерактивные макросы
  блокируются preflight'ом (`INTERACTIVE_UI_NOT_SUPPORTED`).
- Не поддерживает облачные сценарии и Office 365 (работа только с настольным Excel через COM).
- Не поддерживает не-Windows платформы (COM и VBOM доступны только на Windows).

Полный перечень ограничений — в [Возможности Excel](06-excel-capabilities.md).

## Основные зависимости

| Компонент | Требование |
|-----------|------------|
| ОС | Windows 10/11 x64 (валидировано на Windows 10 Enterprise LTSC 2019, сборка 17763) |
| Excel | Настольный Microsoft Excel с COM/VBE (валидировано на 16.0.14332 x64) |
| Python | `>=3.11,<3.15` (валидировано на 3.14 x64) |
| Менеджер пакетов | [uv](https://docs.astral.sh/uv/) |
| `mcp` | `>=2.3,<3` (зафиксировано 2.3.0) |
| `pydantic` | `>=2.0.0` (зафиксировано 2.14.4) |
| `pywin32` | `>=306` (зафиксировано 312) |
| `comtypes` | `>=1.2.0` (зафиксировано 1.5.0) |
| Настройка Excel | Trust Center: «Trust access to the VBA project object model» (VBOM), задаётся вручную на каждой машине |

Процедура установки и проверки описана в [Установка и настройка](02-installation-and-setup.md),
переменные окружения — в [Конфигурация](03-configuration.md).

## Документация

- [Установка и настройка](02-installation-and-setup.md)
- [Конфигурация](03-configuration.md)
- [Архитектура](04-architecture.md)
- [Справочник инструментов MCP](05-mcp-tools-reference.md)
- [Возможности Excel](06-excel-capabilities.md)
- [Рабочие процессы и примеры](07-workflows-and-examples.md)
- [Безопасность](08-security.md)
- [Ограничения и совместимость](09-limitations-and-compatibility.md)
- [Устранение неполадок](10-troubleshooting.md)
- [Разработка и поддержка](11-development-and-maintenance.md)
- [Корневой README](../../README.md)