---
id: ru-xlvba-workflows-and-examples
domain: ru-07-workflows-and-examples
tags:
  - workflows
  - examples
  - mcp
  - excel
  - vba
related:
  - ru-xlvba-overview
  - ru-xlvba-excel-capabilities
  - ru-xlvba-mcp-tools-reference
  - ru-xlvba-configuration
  - ru-xlvba-security
  - ru-xlvba-troubleshooting
---

## Назначение

Документ показывает **реальные, применимые рабочие процессы** для MCP-сервера `xlvba`. Каждый
рабочий процесс перечисляет свои предпосылки, точные вызовы инструментов с JSON-аргументами,
ожидаемый результат и ограничения.

Каждый пример использует опубликованные имена MCP-инструментов и имена параметров из
[Справочник инструментов MCP](05-mcp-tools-reference.md). Ни один из примеров не выполнялся при
написании этого документа: это рецепты, а не транскрипты.

Возможности, лежащие в основе этих рабочих процессов, классифицированы в
[Возможности Excel](06-excel-capabilities.md). Настройки — в [Конфигурация](03-configuration.md);
модель безопасности и approval — в [Архитектура](04-architecture.md).

## Соглашения

| Соглашение | Значение |
|------------|----------|
| `<workbook_id>` | Назначенный сервером идентификатор открытой рабочей копии, возвращаемый `excel_open`. Это **не** путь к файлу. Заменяйте его в каждом последующем вызове. |
| JSON-блоки | Показывают объект **аргументов** инструмента ровно так, как он отправляется через MCP. Блок — это не полный конверт JSON-RPC. |
| Префикс `xlvba_` | Клиенты показывают опубликованные имена с префиксом сервера. Например, в Kilo `excel_open` вызывается как `xlvba_excel_open`. Этот документ использует опубликованные имена без префикса; добавляйте префикс вашего клиента. |
| **Approval-гейт (режим restricted)** | Вызов является мутирующей или исполняющей операцией. В режиме `full` по умолчанию он авторизован (перед мутацией всё равно нужен базовый снапшот); в режиме `restricted` он отклоняется с `APPROVAL_DENIED`, если нет in-process записи allowlist или живого гранта. См. [Архитектура](04-architecture.md). |
| Только чтение | Без изменения книги; разрешено автоматически политикой approval. |

Пути — это Windows-пути. Путь к книге должен находиться внутри `EXCEL_VBA_MCP_ROOTS` (по умолчанию
текущий рабочий каталог и системный temp-каталог); иначе вызов завершается с `PATH_NOT_ALLOWED`.

## Рабочий процесс 1: Открыть и осмотреть книгу (только чтение)

**Предпосылки.** Сервер запущен, книга находится внутри разрешённого корня.

**Шаг 1 — открыть read-only inspect-сессию.**

```json
{
  "path": "C:/data/report.xlsm",
  "mode": "inspect",
  "read_only": true
}
```

Вызовите `excel_open`. Ожидаемый результат содержит:

```json
{ "workbook_id": "1f0c…", "name": "report.xlsm", "format": "xlsm", "path": "<temp>/excel_vba_mcp/work/1f0c….xlsm", "source_path": "C:/data/report.xlsm", "read_only": true, "has_vba": true, "excel_version": "16.0.14332", "session_id": "9b2a…", "mode": "inspect" }
```

**Шаг 2 — прочитать агрегированную модель (или одну инвентаризацию).**

```json
{ "workbook_id": "<workbook_id>", "include_vba": false }
```

Вызовите `workbook_get_model`. Либо вызовите `sheets_list` с тем же `workbook_id`.

**Шаг 3 — закрыть без сохранения.**

```json
{ "workbook_id": "<workbook_id>", "save_changes": false }
```

Вызовите `excel_close`. Рабочая копия удаляется; оригинал не затрагивается.

**Ожидаемый результат.** `WorkbookModel` (или `list[SheetInfo]`, за которым следует `SessionStatus`
с `state="closed"`).

**Ограничения.**

- `mode="inspect"` отключает макросы для открытия; `read_only=true` делает любую мутацию
  недоступной (`WORKBOOK_READ_ONLY`).
- `excel_open` по умолчанию открывает writable, execute-сессию (`mode="execute"`,
  `read_only=False`); передайте `mode="inspect", read_only=true` (как в Шаге 1), чтобы открыть для
  чтения.
- `workbook_get_model` никогда не читает все ячейки и не выполняет VBA.

## Рабочий процесс 2: Прочитать диапазон и его формулы

**Предпосылки.** Книга открыта, как в Рабочем процессе 1.

**Шаг 1 — прочитать типизированные значения.**

```json
{ "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:C10", "offset": 0, "limit": 10 }
```

Вызовите `range_read_values`. Каждая ячейка типизирована (`string`, `number`, `boolean`, `date`,
`error`, `empty`), а ответ несёт `rows`, `cols` и `truncated`.

**Шаг 2 — прочитать формулы в том же окне.** Используйте тот же объект аргументов, что и в Шаге 1,
и вызовите `range_read_formulas`:

```json
{ "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1:C10", "offset": 0, "limit": 10 }
```

У каждой ячейки есть `formula2` и, при отличии, `formula`. Такие префиксы, как `_xlfn`, `_xlws` и
`_xlpm`, сохраняются и перечисляются в `notes`.

**Ожидаемый результат.** Сетка `RangeValues` и сетка `RangeFormulas` для одного окна.

**Ограничения.**

- Чтения ограничены: 100 000 ячеек на цель и 50 000 ячеек на ответ. Читайте бо́льшую область
  порциями через `offset`/`limit`.
- `range_read_values` возвращает значения, которые Excel удерживает в данный момент. Он не форсирует
  пересчёт и не вычисляет формулы.

## Рабочий процесс 3: Осмотреть VBA и экспортировать проект

**Предпосылки.** Книга открыта для чтения; Excel Trust Center должен доверять доступ к объектной
модели проекта VBA (VBOM), и проект не должен быть заблокирован.

**Шаг 1 — сводка проекта.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `vba_project_info`. Результат сообщает имя проекта, состояние защиты/блокировки, инвентарь
компонентов и ссылки.

**Шаг 2 — инвентарь компонентов.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `vba_list_components`. Каждая запись имеет тип (`standard_module`, `class_module`,
`user_form`, `document_module`), количество строк и нормализованный хэш исходника.

**Шаг 3 — прочитать один компонент.**

```json
{ "workbook_id": "<workbook_id>", "component": "Module1" }
```

Вызовите `vba_read_component`. Сохраните возвращённый `source_hash` — это токен ревизии, требуемый
защищёнными инструментами редактирования VBA.

**Шаг 4 — экспортировать весь проект.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `vba_export_project`. Без `out_dir` экспорт идёт в свежий временный каталог, а манифест
возвращает абсолютные пути. Пользовательский `out_dir` должен находиться внутри разрешённого
артефакт-корня `snapshots` или `artifacts` — по умолчанию `<temp>/excel_vba_mcp/artifacts`
(переопределяется через `EXCEL_VBA_MCP_ARTIFACT_DIR`). Если `out_dir` вне этого корня, вызов
завершается отказом с `PATH_NOT_ALLOWED`; в `context.allowed_root` (и `context.expected_root`)
указывается **абсолютный** путь разрешённого артефакт-корня, чтобы можно было повторить запрос с
корректным путём (символьное имя `"artifacts"` больше не возвращается). Результат — манифест с
хэшами по файлам.

**Ожидаемый результат.** `VbaProjectInfo`, список компонентов, один `VbaComponentSource` и
`VbaExportManifest`.

**Ограничения.**

- Без доверия VBOM инструменты возвращают `ACCESS_VBOM_DISABLED`; заблокированный проект возвращает
  `VBA_PROJECT_LOCKED`.
- Экспорт никогда не изменяет книгу и никогда не удаляет Document-модули.

## Рабочий процесс 4: Снапшот перед изменением

**Предпосылки.** Книга открыта (для `snapshot_create` достаточно read-only).

**Шаг 1 — создать иммутабельный снапшот.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `snapshot_create`. Открытые книги фиксируются через `Workbook.SaveCopyAs`; закрытые книги
копируются. Результат — манифест, содержащий `snapshot_id`, `created_at`, `source_path`,
`stored_path`, `sha256`, `excel_version`, `has_vba` и `model_version`.

**Шаг 2 — перечислить снапшоты.**

```json
{}
```

Вызовите `snapshot_list` (без параметров).

**Шаг 3 — сравнить два снапшота.**

```json
{
  "snapshot_id_a": "snap-…-a",
  "snapshot_id_b": "snap-…-b"
}
```

Вызовите `snapshot_diff`. Результат сообщает `file_identical`, `added`, `removed`, `changed` и
`notes`. Он сравнивает только метаданные манифеста и хэши файлов, но никогда не бинарный diff.

**Ожидаемый результат.** Манифест снапшота, список манифестов и `SnapshotDiff`.

**Ограничения.**

- `snapshot_restore` — **мутирующая** операция (авторизована в режиме `full`; `APPROVAL_DENIED` в
  режиме `restricted`). Она закрывает текущую копию без сохранения и повторно открывает baseline;
  внешние эффекты макросов не отменяются.
- Снапшоты иммутабельны; инструмента удаления нет.

## Рабочий процесс 5: Осмотреть таблицы и прочитать строки

**Предпосылки.** Книга открыта, как в Рабочем процессе 1.

**Шаг 1 — перечислить все таблицы Excel.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `tables_list`. Каждый `TableInfo` сообщает лист, диапазон, диапазоны заголовка/тела/итогов,
стиль, количество строк и колонки.

**Шаг 2 — прочитать одну таблицу.**

```json
{ "workbook_id": "<workbook_id>", "name": "Sales" }
```

Вызовите `table_get` с именем или отображаемым именем таблицы.

**Шаг 3 — прочитать ограниченный срез строк данных.**

```json
{ "workbook_id": "<workbook_id>", "name": "Sales", "offset": 0, "limit": 50 }
```

Вызовите `table_read_rows`. `offset` — 0-based индекс строки данных; `limit` ограничивает
возвращаемые строки. Значения типизированы, а ответ имеет флаг `truncated`.

**Ожидаемый результат.** Инвентарь таблиц, один `TableInfo` и payload `TableRows`.

**Ограничения.**

- Списки полные и ограниченные; пагинации нет.
- `table_read_rows` дополнительно ограничен 2 000 ячеек на ответ.
- Мутации строк авторизованы в режиме `full` (`APPROVAL_DENIED` в режиме `restricted`); см.
  Рабочий процесс 7.

## Рабочий процесс 6: Осмотреть PivotTables

**Предпосылки.** Книга открыта для чтения. Обновляемы только PivotTables на базе базы
данных/таблицы; внешние, OLAP и Data Model источники доступны только для чтения.

**Шаг 1 — инвентаризировать PivotTables и кэши.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `pivot_tables_list`. Инвентаризация — обход по листам; общие кэши сообщаются один раз с
затронутыми таблицами.

**Шаг 2 — прочитать одну PivotTable.**

```json
{ "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1" }
```

Вызовите `pivot_table_get`. Результат включает `TableRange1`/`TableRange2`, диапазон тела данных,
макет полей row/column/page/hidden/data и тип источника.

**Шаг 3 — перечислить её поля.**

```json
{ "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1" }
```

Вызовите `pivot_fields_list`. Каждое поле сообщает ориентацию, 1-based позицию внутри ориентации,
агрегацию и состояние ручного фильтра.

**Шаг 4 — прочитать состояние фильтра одного поля.**

```json
{ "workbook_id": "<workbook_id>", "pivot_table": "PivotTable1", "field": "Region" }
```

Вызовите `pivot_field_get_filters`. Результат сообщает `all_items_visible`, видимые и скрытые
`PivotItems`, классифицированный вид фильтра и любые обнаруженные `PivotFilters`.

**Ожидаемый результат.** `PivotInventory`, `PivotTableInfo`, список полей и `PivotFieldFilters`.

**Ограничения.**

- Создание или удаление PivotTables/PivotCaches не поддерживается.
- `PivotFilters.Add` (label/value/date/top-bottom) не поддерживается на валидированной сборке;
  можно применять только ручные фильтры по видимым элементам.
- Обновление, макет, агрегация, свойства и фильтры авторизованы в режиме `full`
  (`APPROVAL_DENIED` в режиме `restricted`).

## Рабочий процесс 7: Мутирующий рабочий процесс (режим полного доступа)

В режиме `full` по умолчанию последовательность мутаций выполняется без грантов на действие:
откройте writable рабочую копию (по умолчанию), создайте базовый снапшот, затем выполните одну или
несколько защищённых мутаций. Каждая защищённая мутация всё равно перечитывает postcondition и
записывает аудит-запись. В режиме `restricted` гейт approval отклоняет каждый мутирующий вызов с
`APPROVAL_DENIED`.

**Шаг 1 — открыть writable рабочую копию (по умолчанию).**

```json
{ "path": "C:/data/report.xlsm" }
```

Вызовите `excel_open` (пропуск `mode`/`read_only` открывает `mode="execute", read_only=False`). В
режиме `full` это авторизовано и создаётся рабочая копия; в режиме `restricted` это `EXECUTING` и
возвращает `APPROVAL_DENIED`.

**Шаг 2 — создать базовый снапшот.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `snapshot_create`. Это approval-гейтированная запись в файловую систему: разрешена
автоматически в режиме `full` по умолчанию, отклоняется с `APPROVAL_DENIED` в режиме `restricted`.

**Шаг 3 — записать значения.**

```json
{
  "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "B2:C3",
  "payload": {
    "rows": 2, "cols": 2,
    "values": [
      [ { "kind": "number", "value": 1 }, { "kind": "string", "value": "North" } ],
      [ { "kind": "number", "value": 2 }, { "kind": "string", "value": "South" } ]
    ]
  }
}
```

Вызовите `range_write_values`. В режиме `full` это завершается успешно с `verified=true` и перечитанным
`postcondition_values`. Строка, начинающаяся с `=`, сохраняется как текст, а не как формула.

**Шаг 4 — добавить рабочий лист.**

```json
{ "workbook_id": "<workbook_id>", "name": "Scratch", "sheet_type": "worksheet" }
```

Вызовите `sheet_add`. Завершается успешно в режиме `full`.

**Шаг 5 — добавить строки таблицы.**

```json
{
  "workbook_id": "<workbook_id>",
  "name": "Sales",
  "rows": [
    { "Region": "North", "Amount": 120 },
    { "Region": "South", "Amount": 80 }
  ]
}
```

Вызовите `table_append_rows`. Завершается успешно в режиме `full`. Таблица должна уже существовать; структура
строк/колонок таблицы сохраняется.

**Ожидаемый результат (режим полного доступа).** Каждая мутация возвращает результат с
`verified=true` и перечитанным postcondition (`postcondition_values`, `result_format`,
`affected_rows` и так далее).

**Сохранение.** Правки живут только в рабочей копии; оригинал никогда не записывается. Чтобы
сохранить, вызовите `excel_close` с `save_changes=true` (записывает рабочую копию) или `excel_save`.
Поскольку `excel_close` удаляет рабочую копию, а поздний `excel_open` повторно копирует нетронутый
оригинал, буквальное «закрыть, затем открыть снова» **не** показывает прежних правок. Чтобы
вернуться к прежнему состоянию, используйте `snapshot_create` → `excel_close` → `snapshot_restore`.

**Режим restricted / конфигурация встраивателя.** В режиме `restricted` сервер строит
`ApprovalPolicy(mode="restricted")` с пустым allowlist и не предоставляет инструмента гранта или
allowlist, поэтому каждый вызов `MUTATING`/`EXECUTING` отклоняется с `APPROVAL_DENIED`. Внутрипроцессный
встраиватель, которому нужны мутации в режиме restricted, конструирует сервисы с записью allowlist в
`ApprovalPolicy` или живым грантом, совпадающим с каноническим путём книги (или SHA-256), именем
действия и, для `macro_run`, процедурой. Реализация политики — `src/excel_vba_mcp/security/approval.py`.

**Ограничения.**

- Каждая мутация требует предварительного успешного снапшота (иначе `SNAPSHOT_FAILED`).
- `range_write_values` для диапазона, который меняет структуру строк/колонок ListObject,
  отклоняется с `LISTOBJECT_INTERSECTION`.
- `range_merge` может отбросить неякорные данные, если не передан `allow_data_loss=true` после
  просмотра предпросмотра.

## Рабочий процесс 7b: Запустить макрос в режиме полного доступа (классифицировать → запустить → проверить)

`macro_run` требует execute-режима, writable-сессии и классификации `background_safe`, привязанной
к текущему хэшу книги и ревизии исходников VBA. В режиме `full` по умолчанию грант approval на
запуск не нужен.

**Шаг 1 — открыть (по умолчанию execute/writable) и сделать снапшот.**

```json
{ "path": "C:/data/report.xlsm" }
```

Вызовите `excel_open`, затем `snapshot_create` с возвращённым `workbook_id`.

**Шаг 2 — перечислить макросы и классифицировать цель как background-safe.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `macro_list`, чтобы подтвердить точное имя `Module.Procedure` и текущее состояние
`background_safe`. Затем классифицируйте:

```json
{ "workbook_id": "<workbook_id>", "procedure": "Module1.Stamp", "background_safe": true }
```

Вызовите `macro_classify_background`. Интерактивная/эвристическая процедура отклоняется с
`INTERACTIVE_UI_NOT_SUPPORTED`; устаревший `expected_project_source_hash` отклоняется с
`STALE_TARGET`.

**Шаг 3 — запустить макрос, затем проверить ячейку, которую он записывает.**

```json
{ "workbook_id": "<workbook_id>", "macro_name": "Module1.Stamp", "args": [] }
```

Вызовите `macro_run`. Затем прочитайте ячейку, которую макрос, как ожидается, изменит:

```json
{ "workbook_id": "<workbook_id>", "sheet": "Sheet1", "address": "A1", "offset": 0, "limit": 1 }
```

Вызовите `range_read_values` и подтвердите ожидаемое значение.

**Шаг 4 — снова отредактировать, затем запустить второй макрос (edit→macro→edit→macro).**

Запишите новое значение через `range_write_values` (или `table_update_rows`), затем снова вызовите
`macro_run` для второй процедуры. Каждый защищённый шаг независим; последовательность — ровно то, что
выражает декларативный `test_run_suite` с шагами `macro_run`, несущими необязательное поле
`background_safe`.

**Ожидаемый результат.** Каждый `macro_run` возвращает `success` (или `blocked` с
`BACKGROUND_SAFETY_NOT_VERIFIED`, если классификация была инвалидирована изменением исходников VBA),
а каждая правка возвращает проверенный postcondition.

**Ограничения.**

- Если исходники VBA изменяются между классификацией и запуском, классификация инвалидируется, и
  макрос оказывается `blocked` с `BACKGROUND_SAFETY_NOT_VERIFIED`; переклассифицируйте после
  перечитывания.
- Токен-скан R13 — эвристический запрет, а не статическое доказательство: динамическая косвенность
  `CallByName`/`ExecuteGlobal`/`Eval` и транзитивный UI покрыты не полностью. Список запретов всё
  же включает `AppActivate`, `Application.Visible`, `.Activate` и `WindowState`, поэтому прямое
  использование перехватывается; UI, достижимый только косвенно, всё ещё может обойти подстроковый
  скан.
- Внешние эффекты макросов никогда не откатываются снапшотом.

## Рабочий процесс 8: Декларативный тестовый набор

**Предпосылки.** Открытая книга (для открытия достаточно read-only-осмотра). Набор открывает свои
собственные одноразовые копии в execute-режиме; в режиме `full` по умолчанию это авторизовано,
тогда как `restricted` возвращает прогон как `blocked` (`APPROVAL_DENIED`). Шаг `macro_run` по
неклассифицированной процедуре — `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`) в любом режиме.

**Шаг 1 — отправить набор с закрытым словарём.**

```json
{
  "suite": {
    "suite": "smoke",
    "workbook_id": "<workbook_id>",
    "baseline": "create",
    "cases": [
      {
        "id": "sheet-and-range",
        "steps": [
          { "action": "sheet_add", "sheet": "Scratch" },
          {
            "action": "range_write_values",
            "sheet": "Scratch",
            "range_address": "A1:B2",
            "values_grid": {
              "rows": 2,
              "cols": 2,
              "values": [
                [ { "kind": "string", "value": "Name" }, { "kind": "string", "value": "Qty" } ],
                [ { "kind": "string", "value": "Widget" }, { "kind": "number", "value": 7 } ]
              ]
            }
          }
        ],
        "assertions": [
          { "type": "sheet_exists", "sheet": "Scratch" },
          { "type": "sheet_visibility", "sheet": "Scratch", "expected": "visible" },
          { "type": "cell_value_equals", "sheet": "Scratch", "cell": "B2", "expected": 7 }
        ]
      },
      {
        "id": "table-rows",
        "steps": [
          { "action": "table_append_rows", "table": "Sales", "rows": [ { "Region": "North", "Amount": 120 } ] }
        ],
        "assertions": [
          { "type": "table_exists", "table": "Sales" },
          { "type": "table_row_count", "table": "Sales", "value": 4 }
        ]
      },
      {
        "id": "compile-and-macro",
        "steps": [
          { "action": "vba_compile" },
          { "action": "macro_run", "macro": "Module1.Hello", "args": [], "background_safe": true }
        ],
        "assertions": [
          { "type": "compile_status", "expected": "indeterminate" },
          { "type": "macro_status", "expected": "success" },
          { "type": "read_cell", "sheet": "Sheet1", "cell": "A1", "expected": "ok" }
        ]
      }
    ]
  }
}
```

Вызовите `test_suite_run`, передав этот объект как аргумент `suite` (`test_run_suite` эквивалентен и
принимает ту же форму). Словари закрыты:

Закрытые словари содержат **37 действий** и **29 проверок**, определённых в
`src/excel_vba_mcp/domain/test_spec.py` и суммированных в
[Справочник инструментов MCP](05-mcp-tools-reference.md). Приведённые выше примеры используют
подмножество: `sheet_add`, `range_write_values`, `table_append_rows`, `vba_compile`, `macro_run`,
`sheet_exists`, `sheet_visibility`, `cell_value_equals`, `table_exists`, `table_row_count`,
`compile_status`, `macro_status`, `read_cell`.

**Шаг 2 — прочитать сохранённый прогон.**

```json
{ "run_id": "<test_run_id>" }
```

Вызовите `test_get_run` с `test_run_id` из `TestRunResult`. Он возвращает per-case и per-step
статусы, не обращаясь к Excel.

**Ожидаемый результат.** `TestRunResult` с одной записью на кейс и шаг плюс сохранённый прогон под
`runs/`.

**Ограничения.**

- `vba_replace_component_source` и `vba_insert_code` существуют в enum действий, но **отклоняются
  раннером**: декларативный шаг не может нести токен ревизии `expected_source_hash`. Используйте
  защищённые MCP-инструменты для этих правок.
- `macro_status` и `compile_status` проверяют последний запуск макроса кейса и пробу компиляции.
- Неудачный шаг `vba_compile` блокирует последующие шаги `macro_run` в том же кейсе; `indeterminate`-компиляция
  не является провалом.
- Шаг `macro_run` может нести необязательное поле `background_safe` только с данными; раннер
  классифицирует процедуру через защищённый путь `classify_background` перед диспетчеризацией.
- В режиме `restricted` каждый кейс — `blocked` (`APPROVAL_DENIED`); шаг неклассифицированного
  макроса — `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`) в любом режиме.

## Рабочий процесс 9: Диагностировать состояние компиляции

**Предпосылки.** Книга открыта для чтения; для любых доказательств VBA требуется доверие VBOM.

**Шаг 1 — запросить результат компиляции.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `vba_compile`. На валидированной сборке результат fail-closed:

```json
{
  "workbook_id": "1f0c…", "status": "indeterminate",
  "reason": "NON_ACTIVATING_COMPILE_UNAVAILABLE",
  "diagnostics": { "is_compile": false, "vbom_available": true, "protection": 0, "locked": false, "component_count": 3, "broken_references": [] },
  "command_discovery": { "route_usable": false },
  "notes": ["NOT a compile: non-activating VBA compilation is unavailable on this build.", "status=indeterminate reason=NON_ACTIVATING_COMPILE_UNAVAILABLE"]
}
```

**Шаг 2 — перечитать кэшированную диагностику.**

```json
{ "workbook_id": "<workbook_id>" }
```

Вызовите `vba_get_compile_diagnostics`. Он возвращает те же диагностические данные без повторного
сбора.

**Ожидаемый результат.** Статус `indeterminate` со структурированной причиной и частичной
диагностикой.

**Ограничения.**

- `vba_compile` никогда не возвращает `compiled` или `compile_error`; `diagnostics.is_compile`
  всегда `false`. Чистый результат `indeterminate` **не** является доказательством того, что проект
  компилируется.
- Возможные причины: `NON_ACTIVATING_COMPILE_UNAVAILABLE`, `ACCESS_VBOM_DISABLED`,
  `VBA_PROJECT_LOCKED`.
- UI не показывается и не активируется.

## Рабочий процесс 10: Обрабатывать ошибки

Сбои инструментов всплывают как структурированная ошибка MCP-инструмента, сообщение которой — это
JSON `to_dict()`: стабильный `code` и объект `context` (например,
`{"code": "APPROVAL_DENIED", "context": {…}}`). Та же ошибка логируется в `stderr` и записывается в
аудит-след. Таблица ниже показывает коды, которые вероятнее всего появятся, и правильную реакцию.

| Код | Где всплывает | Типичная причина | Что делать |
|-----|---------------|------------------|------------|
| `PATH_NOT_ALLOWED` | `excel_open`, `vba_export_project` | Путь вне `EXCEL_VBA_MCP_ROOTS`, device/UNC-путь, traversal, побег через symlink или запрещённое расширение. `context.reason` называет категорию (`outside_roots`, `extension_not_allowed`, `unc_denied`, `symlink_escape`, `root_mismatch`, `traversal`, `empty_path`, `device_path_denied`); `context.allowed_roots` перечисляет загруженные корни. | Поместите книгу внутрь разрешённого корня или расширьте `EXCEL_VBA_MCP_ROOTS` и перезапустите сервер; используйте обычный путь и разрешённое расширение. См. [Конфигурация](03-configuration.md). |
| `LISTOBJECT_INTERSECTION` | `range_write_values`, `range_write_formulas`, `range_copy`, `range_insert`, `range_delete`, `range_merge`, `range_unmerge` | Цель пересекает таблицу Excel (ListObject). `context.tables` перечисляет пересекающиеся таблицы. | Используйте инструменты Tables (`table_append_rows`, `table_update_rows`, `table_add_column` и так далее) вместо структурных операций диапазона. |
| `ACCESS_VBOM_DISABLED` | `vba_*`, `macro_list`, `vba_compile` | Excel не может читать объектную модель проекта VBA; доверие VBOM в Trust Center выключено. | Включите **Trust access to the VBA project object model** в Excel Trust Center → Macro Settings, закройте Excel и повторите. Корпоративная политика может это блокировать. |
| `APPROVAL_DENIED` | `excel_open` (writable/execute), любой мутирующий инструмент, `macro_run`, `test_run_suite` | Сервер работает в режиме `restricted` (`EXCEL_VBA_MCP_ACCESS_MODE=restricted`), и ни одна запись allowlist или живой грант не совпадает. | Сбросьте `EXCEL_VBA_MCP_ACCESS_MODE` (или задайте `full`) для поведения полного доступа по умолчанию; сервер в режиме restricted предоставляет через MCP только работу для чтения. |
| `BACKGROUND_SAFETY_NOT_VERIFIED` | `macro_run` (в `MacroRunResult.error_code`, статус `blocked`) | Процедура не классифицирована `background_safe` для текущего хэша книги и ревизии исходников VBA. | Макрос **не** был запущен. Классифицируйте его через `macro_classify_background` (переклассифицируйте после любого изменения исходников VBA) и повторите. Не повторяйте слепо. |
| `INTERACTIVE_UI_NOT_SUPPORTED` | `macro_run`, `macro_classify_background`, `test_run_suite` | Процедура ссылается на интерактивный UI, эвристична или была классифицирована `background_safe`, будучи небезопасной. | Осмотрите процедуру и удалите интерактивный UI либо оставьте её неклассифицированной. Не обходите preflight. |
| `COM_BUSY` | Любой мутирующий инструмент; `vba_replace_component_source`/`vba_insert_code` | Реальная конкуренция за COM (сервер занят) или конфликт ревизий `source_hash` VBA (`context.reason="revision_conflict"`). Отказ approval — отдельный код (`APPROVAL_DENIED`). | Прочитайте `context.reason`. Повторите после завершения операции; при конфликте ревизий перечитайте и передайте текущий `expected_source_hash`. |
| `WORKBOOK_READ_ONLY` | Любой мутирующий инструмент | Сессия была открыта с `read_only=true`. | Откройте заново для записи (`excel_open` без `read_only`) или выполняйте только работу для чтения. |
| `SNAPSHOT_FAILED` | Любой мутирующий инструмент | Для книги нет успешного снапшота. | Сначала вызовите `snapshot_create`. |
| `UNSUPPORTED_PIVOT_SOURCE` | `pivot_table_refresh`, `pivot_cache_refresh`, `pivot_field_*` | Источник кэша — внешний, OLAP или Data Model. | Обновление или перенастройка для такого источника невозможны; read-only-осмотр по-прежнему работает. |
| `AMBIGUOUS_FILTER` | `table_update_rows`, `table_delete_rows` | Фильтр `where` совпал с нулём или несколькими строками. | Используйте явные `indices` или уточните фильтр, пока он не совпадёт ровно с одной строкой. |
| `INVALID_RANGE` | Инструменты диапазонов | Некорректный или сверхбольшой адрес, неверная форма payload или объединение, заблокированное потерей данных. | Исправьте адрес/payload; передавайте `allow_data_loss=true` только после просмотра предпросмотра объединения. |

Дополнительные коды, объявленные сервером, включают `VBA_PROJECT_LOCKED`, `MISSING_REFERENCE`,
`RPC_E_CALL_REJECTED`, `COMPILE_ERROR`, `COMPILE_INDETERMINATE`, `MACRO_BLOCKED`,
`MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `EXCEL_HUNG`, `PROTECTED_TARGET`,
`STALE_TARGET`, `REFRESH_FAILED` и `MODAL_UI_BLOCKED`. Полный контракт — в
[Справочник инструментов MCP](05-mcp-tools-reference.md).

## Безопасное комбинирование инструментов

### Правила порядка

1. **Сначала открыть.** Каждому инструменту, принимающему `workbook_id`, нужна открытая рабочая
   копия. Начните с `excel_open` и сохраните возвращённый `workbook_id`.
2. **Осмотр через `mode="inspect", read_only=true`.** Используйте это для каждого read-only-рабочего
   процесса. Это отключает макросы и делает случайную мутацию невозможной.
3. **Снапшот перед мутацией.** Вызовите `snapshot_create` после открытия для записи и до первого
   мутирующего вызова. Без него каждая мутация завершается с `SNAPSHOT_FAILED`.
4. **Читайте, прежде чем писать.** Используйте `vba_read_component` для `source_hash`, `table_get`
   для точного имени таблицы, `pivot_table_get`/`pivot_fields_list` для имён PivotTable и полей и
   `names_list` перед тем, как трогать определённое имя.
5. **Один объект на мутацию.** Меняйте по одной таблице, диапазону, листу или полю за раз, затем
   осматривайте возвращённый postcondition (`verified`, `postcondition_values`, `result_format`,
   `affected_rows`).
6. **Перечитывайте после мутации.** Не предполагайте успех по отсутствию ошибки; используйте
   возвращённый postcondition или свежее чтение.
7. **Закрывайте осознанно.** `excel_close` без `save_changes` — безопасный выбор по умолчанию.
   Сохранение — отдельное, защищённое решение.
8. **Никогда не повторяйте заблокированный вызов слепо.** Заблокированный макрос или отклонённая
   мутация будут заблокированы снова. Прочитайте `context.reason` и измените предусловие вместо
   этого.

### Цепочка защиты, которую нужно держать в уме

| Порядок | Гейт | Код ошибки |
|---------|------|------------|
| 1 | Writable-сессия (мутирующие инструменты) | `WORKBOOK_READ_ONLY` |
| 2 | Базовый снапшот | `SNAPSHOT_FAILED` |
| 3 | Approval (`full` по умолчанию / `restricted`) | `APPROVAL_DENIED` (режим restricted) |
| 4 | Валидация, специфичная для операции (форма payload, явный `shift`, однозначный `where`, отсутствие пересечения с ListObject, поддерживаемый источник сводной) | `INVALID_RANGE`, `AMBIGUOUS_FILTER`, `LISTOBJECT_INTERSECTION`, `UNSUPPORTED_PIVOT_SOURCE` |
| 5 | Перечитывание postcondition + аудит | `verified=false` в результате |

### Что делать и чего не делать

| Делать | Не делать |
|--------|-----------|
| Использовать read-only-сессии для осмотра и отчётности. | Предполагать, что каждая сессия должна быть read-only. |
| Открывать для записи (`mode="execute"`, по умолчанию) для мутаций в режиме full. | Предполагать, что мутациям нужен грант в режиме full. |
| Классифицировать через `macro_classify_background` перед `macro_run`. | Запускать макрос без классификации `background_safe`. |
| Делать снапшот перед любым мутирующим рабочим процессом. | Пропускать снапшот перед мутацией. |
| Использовать инструменты Tables для структуры таблицы. | Использовать `range_insert`/`range_delete` над ListObject. |
| Сохранять `expected_source_hash` из `vba_read_component`. | Редактировать VBA без токена ревизии. |
| Считать `indeterminate`-компиляцию «неизвестно», а не «пройдено». | Считать чистый `vba_compile` доказательством компиляции. |
| Читать `context.reason` при каждом отказе. | Слепо повторять заблокированный макрос или отклонённую мутацию. |

Связанные документы: [Обзор](01-overview.md), [Возможности Excel](06-excel-capabilities.md),
[Справочник инструментов MCP](05-mcp-tools-reference.md), [Конфигурация](03-configuration.md),
[Архитектура](04-architecture.md), [Безопасность](08-security.md),
[Устранение неполадок](10-troubleshooting.md) и [корневой README](../../README.md).