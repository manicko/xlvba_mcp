---
id: ru-xlvba-mcp-tools-reference
domain: ru-05-mcp-tools-reference
tags:
  - mcp
  - tools
  - reference
  - excel
  - vba
related:
  - ru-xlvba-overview
  - ru-xlvba-architecture
  - ru-xlvba-excel-capabilities
  - ru-xlvba-security
  - ru-xlvba-limitations-and-compatibility
  - ru-xlvba-troubleshooting
---

## Назначение

Справочник описывает **41 инструмент** MCP-сервера `xlvba`: назначение, значения `action`,
параметры, возвращаемое значение, побочные эффекты, типовые ошибки и правдивые аннотации. Документ
основан на фактических схемах инструментов, полученных с живого `MCPServer`, и на исходных модулях
`api/*`, `application/*` и `adapters/*`.

Поверхность инструментов была консолидирована из прежнего каталога из **92 инструментов**
(промежуточный счёт «91 опубликовано» в старых редакциях был сведён к 92, как только был учтён
`server_health`) до финальных **41 инструментов** за один переход (single switchover).
Недеструктивные глаголы по доменам были объединены в инструменты с диспетчеризацией по `action`
(с `_read`-компаньонами там, где есть чтения); разрушительные, высокорисковые и lifecycle-операции
оставлены отдельными инструментами с неизменными именами. См.
[Миграция (single switchover)](02-installation-and-setup.md#миграция-поверхности-инструментов-single-switchover)
для заметки о миграции клиента.

Документ не пересказывает архитектуру — она описана в [Архитектура](04-architecture.md).
Настройки — в [Конфигурация](03-configuration.md), модель безопасности — в
[Архитектура](04-architecture.md), пределы — в [Возможности Excel](06-excel-capabilities.md).

## Как читать справочник

| Элемент | Значение |
|---------|----------|
| `Инструмент` | Точное MCP-имя. Клиенты показывают его с префиксом (например, Kilo — `xlvba_workbook`). В таблицах — без префикса. |
| `Действия` | Значения `action`, по которым диспетчеризуется инструмент. У отдельного (standalone) инструмента единственное неявное действие равно его имени. |
| `Параметры` | `имя:тип`; `?` означает необязательный параметр; `=value` — значение по умолчанию. `—` означает отсутствие параметров. Селектор `action` и поля, специфичные для действия, проверяются в обработчике (per-action required params). |
| `Возврат` | Тип DTO или структура ответа. Все DTO сериализуются в JSON и не содержат COM-ссылок. |
| `Побочные эффекты` | Что меняется за пределами ответа: файлы, снапшоты, состояние книги, артефакты, аудит. |
| `Ошибки` | Стабильные коды `ErrorCode` или иные коды, которые может вернуть инструмент. |
| `R/D/I` | Правдивые MCP-аннотации: `readOnly` / `destructive` / `idempotent` (`T`/`F`), взятые из живых аннотаций инструментов. |

Общие соглашения:

1. **Режим доступа.** В режиме `full` (по умолчанию, `EXCEL_VBA_MCP_ACCESS_MODE=full`)
   `ApprovalPolicy` авторизует мутирующие и исполняющие операции без гранта на действие. В режиме
   `restricted` действует deny-by-default: такие операции отклоняются с `APPROVAL_DENIED`, если нет
   in-process записи allowlist или живого гранта. Grant-инструмента для approval или редактирования
   allowlist через MCP нет, поэтому сервер в режиме `restricted` через MCP предоставляет только
   работу для чтения. По таксономии ошибок B4: отказ approval — это `APPROVAL_DENIED`;
   `WORKBOOK_READ_ONLY` поднимается только для read-only-сессии; неизвестный или уже не открытый
   `workbook_id` — это `WORKBOOK_NOT_FOUND` (ошибка вызывающей стороны, а не конкуренция);
   устаревший `expected_source_hash` — это `REVISION_CONFLICT`; отсутствующая фигура/элемент по имени
   — `NOT_FOUND` (F6); а `COM_BUSY`/`RPC_E_CALL_REJECTED` зарезервированы только для реальной
   конкуренции за COM. `MACRO_BLOCKED` — только для Office macro-security / не-execute контекстов.
   Потерянный owned-экземпляр Excel на обычных путях чтения/мутации сообщается как
   `WORKBOOK_NOT_FOUND`, а в `macro_run` / `vba_compile_diagnostics` — как `EXCEL_HUNG`
   (`needs_recovery`), где требуется восстановление — кроме `macro_run`, который при
   зависании/таймауте восстанавливает сессию сам (F2b). В таблицах инструментов `COM_BUSY` у
   mutating-строки означает реальную конкуренцию, а не отказ approval, плохой id или конфликт
   ревизий.
2. **Сначала `workbook` `action="open"`.** Любой инструмент, которому нужен `workbook_id`, требует
   предварительного открытия книги. `workbook_id` — это идентификатор рабочей копии, выданный
   сервером, а не путь к файлу.
3. **Снапшот перед мутацией.** Мутирующие инструменты требуют успешного снапшота
   (`snapshot_create`); иначе — `SNAPSHOT_FAILED`.
4. **Границы.** Чтение диапазонов и таблиц ограничено (`offset`/`limit`, флаг `truncated`);
   запросы сверх лимита отклоняются, а не «молча» урезаются.
5. **Никакого GUI.** Инструменты не активируют окна, не показывают диалоги и не отправляют клавиши.

## Аннотации (правдивые подсказки)

Каждый инструмент advertise-ит тройку MCP `readOnlyHint` / `destructiveHint` / `idempotentHint`,
вычисленную из матрицы правды по действиям (семантика MCP по умолчанию: отсутствие не нейтрально).
Подсказки — это **подсказки, а не энфорсмент**: approval в Kilo остаётся name-glob'ом `xlvba_*`
(см. [Конфигурация](03-configuration.md) и заметку о миграции). Распределение по 41 инструменту:

| Аннотация | Кол-во | Инструменты |
|-----------|--------|-------------|
| `readOnly=true` | 13 | `workbook_read`, `excel_health`, `snapshot_list`, `snapshot_diff`, `vba_read`, `job_get`, `sheet_read`, `table_read`, `pivot_read`, `range_read`, `name_read`, `server_health`, `xlvba_environment` |
| `destructive=true` | 16 | `workbook`, `excel_save_as`, `excel_recover`, `snapshot_restore`, `test`, `vba_replace_component_source`, `vba_insert_code`, `macro_run`, `sheet_delete`, `table_delete_rows`, `table_delete_column`, `pivot_field`, `range_structural`, `range_delete`, `range_merge`, `name_delete` |
| `idempotent=true` | 21 | `workbook_read`, `excel_health`, `snapshot_list`, `snapshot_diff`, `vba_read`, `job_get`, `job_request_cancel`, `sheet_read`, `table_read`, `pivot_read`, `pivot_refresh`, `range_read`, `range_merge`, `name`, `name_read`, `shape`, `control`, `server_health`, `xlvba_environment`, `macro`, `pivot_field` |

Заметные, доказанные матрицей случаи:

- **`workbook` — `destructive=true`**, потому что у его действия `close` есть разрушительный путь
  `discard_changes=true`; `close` **по умолчанию отказывает** (D1), а `open`/`save` недеструктивны.
  См. [D1 close disposition](#workbook).
- **`range_structural` НЕ идемпотентен** (H4): повторный `insert` дублирует строки/ячейки, а
  `unmerge` отбрасывает объединения.
- **Инструменты с двойной подсказкой — ровно `range_merge` и `pivot_field`** (`destructive=true`
  **и** `idempotent=true`), оба доказаны матрицей (`pivot_field` через
  `field-set-layout(delete=True)`).
- **`macro` — `readOnly=false, destructive=false, idempotent=true`**: `list`/`get-last-result` —
  чтения, а `classify-background` — идемпотентный bookkeeping; в целом инструмент не read-only.
- **`vba_compile_diagnostics` — `readOnly=false, destructive=false, idempotent=false`**: он **не**
  компилирует (на валидированной сборке всегда `indeterminate`) и немутирующий, но не advertise-ится
  как read-only. Его STA-проба помечается `op_class=EXECUTING` только для повтора/аудита и не
  управляет approval-гейтом.

## Сокращения

| Сокращение | Значение |
|------------|----------|
| `STA` | Single-Threaded Apartment — модель COM-апартамента, в которой живёт Excel. |
| `DTO` | Data Transfer Object — типизированная структура без COM-ссылок. |
| `VBOM` | VBA project object model — объектная модель проекта VBA (требует доверия в Trust Center). |
| `ListObject` | Excel Table — структурированная таблица листа. |
| `BG06` | Правило классификации `background_safe` для фонового выполнения макросов. |
| `R13` | Preflight интерактивного UI (без учёта регистра): `MsgBox`, `InputBox`, `.Show`, `FileDialog`, `Shell`, `CreateObject`, `FollowHyperlink`, `SendKeys`, `DoEvents`, `GetOpenFilename`, `GetSaveAsFilename`, `CallByName`, `ExecuteGlobal`, `Execute`, `Eval`, `OnTime`, `DisplayAlerts`, `.Run`, `CommandBars`, `AppActivate`, `Application.Visible`, `.Activate`, `WindowState`. |
| `VBE` | VBA Editor — среда разработки VBA. |
| `UIA` | UI Automation — доступ к дереву элементов Windows (только чтение, без активации). |

## Сводка по группам

| Группа | Кол-во | Инструменты |
|--------|--------|-------------|
| Lifecycle / сессия | 5 | `workbook`, `workbook_read`, `excel_save_as`, `excel_health`, `excel_recover` |
| Снапшоты | 4 | `snapshot_create`, `snapshot_list`, `snapshot_restore`, `snapshot_diff` |
| Jobs | 2 | `job_get`, `job_request_cancel` |
| VBA | 5 | `vba_read`, `vba_export_project`, `vba_replace_component_source`, `vba_insert_code`, `vba_compile_diagnostics` |
| Macro / test | 3 | `macro`, `macro_run`, `test` |
| Листы / таблицы / имена | 10 | `sheet`, `sheet_read`, `sheet_delete`, `table`, `table_read`, `table_delete_rows`, `table_delete_column`, `name`, `name_read`, `name_delete` |
| Диапазоны | 5 | `range_read`, `range_write`, `range_structural`, `range_delete`, `range_merge` |
| Pivot | 3 | `pivot_read`, `pivot_refresh`, `pivot_field` |
| Объекты | 2 | `shape`, `control` |
| Возможности / здоровье | 2 | `xlvba_environment`, `server_health` |
| **Итого** | **41** | — |

Сумма по группам: 5 + 4 + 2 + 5 + 3 + 10 + 5 + 3 + 2 + 2 = **41**.

## Lifecycle / сессия

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `workbook` | `open`, `close`, `save` | Открывает книгу как скрытую рабочую копию; закрывает; сохраняет. | `action`, `path?`, `mode?=execute`, `read_only?=false`, `background?=false`, `workbook_id?`, `save_changes?`, `discard_changes?` | `WorkbookInfo` / `SessionStatus` / `RunningJob` | Копия файла в `work/`; close удаляет копию; save записывает рабочую копию; аудит | `PATH_NOT_ALLOWED`, `APPROVAL_DENIED`, `EXCEL_HUNG`, `CONFIRMATION_REQUIRED` | F/T/F |
| `workbook_read` | `status`, `list`, `model` | Состояние сессии/книги; список открытых книг; агрегированная модель книги. | `action`, `workbook_id?`, `include_vba?=true` | `SessionStatus` / `list[WorkbookInfo]` / `WorkbookModel` | — | — (состояние VBA — в `vba_access`) | T/F/T |
| `excel_save_as` | (standalone) | Сохраняет рабочую копию по (возможно новому) целевому пути без перепривязки сессии. | `workbook_id`, `path`, `overwrite_existing?=false`, `overwrite_source?=false` | `SaveAsResult` | Сброс рабочей копии (`wb.Save`), затем атомарное temp→replace копирование + один `<target>.bak`; аудит | `PATH_NOT_ALLOWED`, `PROTECTED_TARGET`, `STALE_TARGET`, `APPROVAL_DENIED` | F/T/F |
| `excel_health` | (standalone) | Состояние воркера/сессии без принудительного перезапуска и без COM. | — | `ExcelHealth` | — | — (всегда payload) | T/F/T |
| `excel_recover` | (standalone) | Agent-directed восстановление owned-процесса Excel (fail-closed; originals-safe). | `reason?=agent_directed`, `confirm?=false` | `RecoveryResult` | Завершение зависшего owned `EXCEL.EXE` через удерживаемый handle; отбрасывание рабочих копий; пересоздание свежего скрытого экземпляра; аудит | `CONFIRMATION_REQUIRED`; иначе сообщает `ok=false` / `recovery_possible=false` при отказе | F/T/F |

### workbook

`workbook` диспетчеризуется по `action`:

- `open` — `mode="execute"` (по умолчанию) открывает writable, execution-capable сессию;
  `mode="inspect"` и/или `read_only=true` открывает сессию с отключёнными макросами или только для
  чтения. Авторизовано в `full`; отклоняется с `APPROVAL_DENIED` в `restricted` для
  writable/execute-открытия.
- `close` — **по умолчанию отказ (D1).** Чистая книга закрывается нормально. При наличии
  несохранённых изменений `close` **по умолчанию отказывает**; отбрасывание требует явного
  `discard_changes=true`. Сохраняется явная опция сохранения при закрытии (`save_changes=true`),
  где поддерживается; **противоречивые флаги** (например, `discard_changes=true` вместе с
  `save_changes=true`) отклоняются со структурированным `ExcelVbaMcpError`.
- `save` — явно сохраняет рабочую копию и возвращает обновлённый `WorkbookInfo` с `dirty=false`.
  Изменения никогда не сохраняются неявно.

`workbook` — `readOnly=false, destructive=true, idempotent=false`: путь `discard_changes=true`
разрушителен, поэтому инструмент несёт destructive-подсказку, хотя `open`/`save` недеструктивны.

### excel_save_as

Сохраняет рабочую копию по новому (или исходному) пути — fail-closed Save As (standalone, own
permission/overwrite policy). `excel_save_as` сбрасывает рабочую копию через `wb.Save` на
STA-воркере, затем копирует её по `path` чисто-файловым атомарным temp→`os.replace` (никогда
`wb.SaveAs`, чтобы сессия оставалась привязана к своей рабочей копии). Один `<target>.bak`
сохраняет любую предыдущую цель при перезаписи.

- **Новый** путь записывается напрямую (`outcome="ok"`).
- **Существующая** цель отклоняется с `PROTECTED_TARGET/target_exists`, если не задан
  `overwrite_existing=true` (который пишет `.bak`).
- **Исходный источник** как цель отклоняется с `PROTECTED_TARGET/source_protected`, если не задан
  `overwrite_source=true`; идентичность источника (SHA-256 + размер), зафиксированная при открытии,
  перепроверяется, поэтому заменённый или изменённый с момента открытия источник даёт
  `STALE_TARGET`.
- **Открытая рабочая копия** как цель всегда отклоняется
  (`PROTECTED_TARGET/target_is_working_copy`).
- `outcome` равен `"ok"` только после финальной проверки размера + SHA-256; прерванная/недоказуемая
  запись сообщает `"unknown"`/`"error"` (никогда success) с `backup_path`, когда `.bak` существует.

### excel_recover

`excel_recover` — agent-directed: allow-by-default в **обоих** режимах (`full` и `restricted`),
аудируется, pid-scoped (никогда не wildcard) и отбрасывает **только рабочие копии** — оригиналы
никогда не изменяются. Он fail-closed: когда не удерживается живой identity-verified handle, он
отклоняется (`confirm` должен быть `true`), а документированный fallback — перезапуск сервера. См.
[Архитектура](04-architecture.md) и [Безопасность](08-security.md).

`ExcelHealth.worker_state` / `session_state` включают `running/deadline_exceeded` (один COM-вызов
превысил свой execution-дедлайн; вызов **не** отменяется и owned pid **не** завершается) — это
отличается от `running/needs_recovery`, который является подлинным recovery-состоянием.

## Снапшоты

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `snapshot_create` | (standalone) | Создаёт иммутабельный снапшот открытой или известной книги. | `workbook_id` | `SnapshotManifest` | Файлы снапшота + `manifest.json`; аудит | `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `EXCEL_HUNG` | F/F/F |
| `snapshot_list` | (standalone) | Список снапшотов и их метаданных. | — | `list[SnapshotManifest]` | — | `SNAPSHOT_FAILED` | T/F/T |
| `snapshot_restore` | (standalone) | Восстанавливает рабочую сессию из снапшота. | `snapshot_id`, `workbook_id`, `confirm?=false` | `SnapshotRestoreResult` | Закрытие текущей копии без сохранения; новая рабочая копия; аудит | `CONFIRMATION_REQUIRED`, `INVALID_RANGE`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `EXCEL_HUNG` | F/T/F |
| `snapshot_diff` | (standalone) | Сравнивает два снапшота по метаданным и хэшу. | `snapshot_id_a`, `snapshot_id_b` | `SnapshotDiff` | — | `SNAPSHOT_FAILED` | T/F/T |

`snapshot_create` и `vba_export_project` записывают файлы и гейтятся approval (F-4): авторизованы
без гранта в `full`, отклоняются с `APPROVAL_DENIED` в `restricted`. `snapshot_restore` не
добавляет второй гейт — его точка approval унаследована от writable-открытия, которое он выполняет
над baseline снапшота. `workbook_id` **обязателен**: он определяет текущую рабочую копию для
отбрасывания, чтобы осталась ровно одна активная копия (отсутствующий/пустой `workbook_id`
отклоняется с `INVALID_RANGE`, `isRetryable=false`). `confirm` **обязателен** (по умолчанию
`false`): отсутствующий или `false` `confirm` отклоняется с `CONFIRMATION_REQUIRED` **до** закрытия
рабочей копии. Ошибка закрытия предыдущей копии **не** ломает восстановление: результат содержит
`previous_close="failed"` с безопасным текстом `previous_close_error`, а предыдущая копия остаётся
открытой (`"closed"` при чистом закрытии, `"not_requested"` если предыдущая копия не задана).
Повторное восстановление снапшота, который уже является активным восстановленным baseline для того
же оригинала, отклоняется (`reason="same_snapshot"`). Снапшот восстанавливает только файлы — он
**не** откатывает внешние эффекты макросов.

**Контракт вызова `snapshot_restore` (R1–R3).** Этот контракт обязателен при скриптовании
восстановления:

- **`workbook_id` обязателен и это СТАРЫЙ id.** Передавайте **предыдущий** id рабочей копии — копию,
  которую нужно закрыть. Отсутствующий/пустой `workbook_id` отклоняется с `INVALID_RANGE`
  (`isRetryable=false`) **до** любого закрытия или открытия.
- **Используйте возвращённый `workbook.id` для всех последующих вызовов.** Восстановление открывает
  **новую** рабочую копию с **новым** id; `SnapshotRestoreResult.workbook.id` — это новый активный
  id. Старый `workbook_id` закрыт и не должен использоваться повторно.
- **`previous_close` сообщает состояние старой копии:** `"closed"` (чистое закрытие), `"failed"`
  (закрытие вызвало ошибку — с безопасным текстом `previous_close_error`) или `"not_requested"`
  (предыдущая копия не задана).
- **Ошибка закрытия предыдущей копии не делает восстановленную копию недействительной.**
  Восстановленная копия уже валидирована и активна; при `"failed"` старая копия просто остаётся
  открытой.
- **Повтор того же снапшота — детерминированная ошибка.** Восстановление снапшота, который уже
  является активным восстановленным baseline для того же оригинала, отклоняется с
  `reason="same_snapshot"` (`isRetryable=false`); восстановите другой снапшот или сначала закройте
  восстановленную копию.
- **Три различных артефакта.** **Оригинал-источник** (`source_path`, никогда не изменяется) — то, к
  чему перепривязывается восстановленная сессия; **иммутабельный снапшот** — файл baseline только
  для чтения; **мутируемая рабочая копия** (`work/<newid>.ext`) — единственное, что редактируется.
  Файл снапшота никогда не является идентичностью сессии.

## Jobs

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `job_get` | (standalone) | Возвращает статус/результат одной фоновой задачи по `job_id` (poll; никогда не трогает COM). | `job_id` | `JobStatus` | — | `NOT_FOUND` | T/F/T |
| `job_request_cancel` | (standalone) | Фиксирует **намерение** отмены задачи (intent-only; никогда не прерывает выполняющийся COM-вызов). | `job_id` | `JobStatus` | — | `NOT_FOUND` | F/F/T |

`job_get` и `job_request_cancel` никогда не трогают COM. `job_request_cancel` только выставляет
`cancel_requested` у задачи, чтобы поллер мог это наблюдать: единственный STA COM-вызов
дорабатывает до конца, и задача всё равно достигает терминального состояния. См.
[Фоновые задачи](#фоновые-задачи).

## VBA

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `vba_read` | `project-info`, `list-components`, `list-procedures`, `read-component`, `get-references`, `search-code`, `get-compile-diagnostics` | Сводка проекта VBA; инвентарь компонентов/процедур/ссылок; исходник модуля; поиск по коду; кэш диагностики компиляции. | `action`, `workbook_id`, `component?`, `pattern?`, `regex?=false` | `VbaProjectInfo` / `list[VbaComponentInfo]` / `list[VbaProcedureInfo]` / `VbaComponentSource` / `list[VbaReferenceInfo]` / `list[VbaSearchMatch]` / `CompileDiagnostics` | — | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `EXCEL_HUNG` | T/F/T |
| `vba_export_project` | (standalone) | Экспортирует все компоненты и возвращает манифест. | `workbook_id`, `out_dir?` | `VbaExportManifest` | Файлы `.bas`/`.cls`/`.frm`/`.frx`; артефакт-корень; аудит | `ACCESS_VBOM_DISABLED`, `PATH_NOT_ALLOWED`, `SNAPSHOT_FAILED` | F/F/F |
| `vba_replace_component_source` | (standalone) | Полностью заменяет модуль Standard/Class. | `workbook_id`, `component`, `new_source`, `expected_source_hash` | `VbaMutationResult` | Изменение проекта VBA; экспорт до/после; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `APPROVAL_DENIED`, `REVISION_CONFLICT`, `PROTECTED_TARGET`, `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | F/T/F |
| `vba_insert_code` | (standalone) | Вставляет код в семантическую позицию модуля Standard/Class. | `workbook_id`, `component`, `code`, `position`, `expected_source_hash` | `VbaMutationResult` | Изменение проекта VBA; аудит | те же; неверная `position` — `COM_BUSY` | F/T/F |
| `vba_compile_diagnostics` | (standalone) | Fail-closed диагностика компиляции; на валидированной сборке всегда `indeterminate`. | `workbook_id` | `CompileResult` | — (не запускает UI) | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED` | F/F/F |

`vba_read` — read-only. `vba_replace_component_source` и `vba_insert_code` работают только с
модулями Standard/Class; Document-модули и UserForms отклоняются с `PROTECTED_TARGET`
(`reason=non_replaceable_component`). Устаревший `expected_source_hash` отклоняется с
`REVISION_CONFLICT` (`reason=revision_conflict`); повторный вызов с тем же хэшем не проходит
revision-guard, поэтому `vba_replace_component_source` **не** идемпотентен.
`vba_compile_diagnostics` никогда не возвращает `compiled` или `compile_error` и не показывает и не
активирует UI.

## Macro / test

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `macro` | `list`, `classify-background`, `get-last-result` | Список вызываемых макросов с `background_safe`; запись классификации процедуры; кэш последнего запуска. | `action`, `workbook_id`, `procedure?`, `background_safe?`, `expected_project_source_hash?` | `list[MacroInfo]` / `BackgroundSafetyClassification` / `MacroRunResult` | `classify-background` меняет in-memory состояние классификации (bookkeeping); аудит | `ACCESS_VBOM_DISABLED`, `VBA_PROJECT_LOCKED`, `MACRO_NOT_FOUND`, `STALE_TARGET`, `INTERACTIVE_UI_NOT_SUPPORTED`, `EXCEL_HUNG` | F/F/T |
| `macro_run` | (standalone) | Запускает один одобренный `background_safe` макрос. | `workbook_id`, `macro_name`, `args?`, `background?=false` | `MacroRunResult` или `RunningJob` | Исполнение VBA; возможные внешние эффекты; аудит | `APPROVAL_DENIED`, `MACRO_BLOCKED`, `MACRO_NOT_FOUND`, `MACRO_RUNTIME_ERROR`, `MACRO_TIMEOUT`, `BACKGROUND_SAFETY_NOT_VERIFIED`, `INTERACTIVE_UI_NOT_SUPPORTED` | F/T/F |
| `test` | `run-suite`, `get-run` | Запускает декларативный тестовый набор (каноническая точка входа) или читает сохранённый прогон. | `action`, `suite?`, `background?=false`, `run_id?` | `TestRunResult` / `RunningJob` | Базовый снапшот, рабочая копия на кейс, артефакты, `runs/`, аудит | `INVALID_RANGE`, `APPROVAL_DENIED`, `SNAPSHOT_FAILED`, `BACKGROUND_SAFETY_NOT_VERIFIED` | F/T/F |

`macro` в целом не read-only (`classify-background` — bookkeeping-запись), но и не разрушителен;
`classify-background` **не гейтится** и не является границей безопасности — `macro_run` сохраняет
approval-гейт и детерминированный вето R13. Классификация привязана к
`(workbook_file_sha256, project_source_hash, procedure)` и инвалидируется при изменении исходников
VBA. `test` — инструмент со смешанным профилем (`run-suite` исполняет, `get-run` читает) и
`destructive=true`, поскольку `run-suite` может диспетчеризовать шаги `macro_run` с внешними
эффектами. `test` принимает `background=true`, чтобы promptly вернуть `RunningJob`.

**Ограничение headless-UI (R13).** Ни `macro_run`, ни декларативный тест-раннер не могут выполнить
процедуру, показывающую UI. Процедура, содержащая `MsgBox`, `InputBox`, `UserForm.Show`, `FileDialog`
или любой иной токен R13, ветируется preflight'ом интерактивного UI **до** диспетчеризации и
возвращает `INTERACTIVE_UI_NOT_SUPPORTED` (макрос не запускается). Это ожидаемое поведение, а не
дефект: сервер владеет скрытым, неактивируемым экземпляром Excel без GUI, поэтому модальное окно
заблокировало бы единственный STA-воркер на неопределённый срок. Вето намеренное и не ослабляется;
удалите интерактивные токены (вне области действия этого сервера) или выполните процедуру в
интерактивной сессии Excel.

## Листы / таблицы / имена

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `sheet` | `add`, `rename`, `copy`, `set-visibility` | Добавляет, переименовывает, копирует лист; задаёт видимость. | `action`, `workbook_id`, `name?`, `new_name?`, `sheet_type?=worksheet`, `visible?`, `target_name?` | `SheetInfo` | Новый/переименованный/скопированный лист; видимость; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `NOT_FOUND`, `INVALID_RANGE`, `COM_BUSY` | F/F/F |
| `sheet_read` | `list`, `get` | Все листы или один лист по имени. | `action`, `workbook_id`, `name?` | `list[SheetInfo]` / `SheetInfo` | — | `NOT_FOUND` | T/F/T |
| `sheet_delete` | (standalone) | Удаляет лист (разрушительно). | `workbook_id`, `name` | `SheetInfo` | Удаление листа; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `NOT_FOUND`, `COM_BUSY` | F/T/F |
| `table` | `create`, `resize`, `append-rows`, `insert-row`, `update-rows`, `add-column` | Создаёт/изменяет размер таблицы; добавляет/вставляет/обновляет строки; добавляет колонку. | `action`, `workbook_id`, `sheet?`, `range_or_name?`, `name?`, `new_range?`, `rows?`, `at?`, `values?`, `updates?`, `header?`, `position?`, `style?` | `TableMutationResult` | Строки/колонки/диапазон таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/F/F |
| `table_read` | `list`, `get`, `read-rows` | Все таблицы, одна таблица или ограниченный срез строк. | `action`, `workbook_id`, `name?`, `offset?=0`, `limit?` | `list[TableInfo]` / `TableInfo` / `TableRows` | — | `COM_BUSY` | T/F/T |
| `table_delete_rows` | (standalone) | Удаляет строки по индексам или однозначному `where` (XOR-guard). | `workbook_id`, `name`, `indices?`, `where?` | `TableMutationResult` | Строки таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/F |
| `table_delete_column` | (standalone) | Удаляет колонку по уникальному заголовку (own permission). | `workbook_id`, `name`, `header` | `TableMutationResult` | Колонка таблицы; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/F |
| `name` | `create`, `update` | Создаёт имя в явной области; обновляет `RefersTo` (область сохраняется). | `action`, `workbook_id`, `name`, `refers_to?`, `scope`, `sheet?` | `DefinedNameInfo` | Новое/обновлённое имя; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/F/T |
| `name_read` | `list`, `get` | Все имена или одно имя по имени и явной области. | `action`, `workbook_id`, `name?`, `scope?`, `sheet?` | `list[DefinedNameInfo]` / `DefinedNameInfo` | — | `COM_BUSY` | T/F/T |
| `name_delete` | (standalone) | Удаляет имя и возвращает его предыдущее значение. | `workbook_id`, `name`, `scope`, `sheet?` | `DefinedNameInfo` | Имя; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/F |

`sheet_delete` отклоняется, если после удаления не останется ни одного видимого листа;
`sheet set-visibility` отклоняет скрытие последнего видимого листа. `table update-rows` принимает
явный индекс или однозначный `where`; `table_delete_rows` требует ровно один из `indices` /
`where` (XOR). `name create` отклоняет дубликат (та же область) или кросс-областный конфликт;
`name update` никогда не меняет область имени.

## Диапазоны

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `range_read` | `get`, `read-values`, `read-formulas`, `get-format`, `get-validation`, `get-conditional-formatting` | Метаданные диапазона; типизированные значения/формулы; формат; проверка; условное форматирование. | `action`, `workbook_id`, `sheet`, `address`, `offset?=0`, `limit?` | `RangeAddress` / `RangeValues` / `RangeFormulas` / `FormatInfo` / `ValidationRule` / `list[ConditionalFormatRule]` | — | `INVALID_RANGE` | T/F/T |
| `range_write` | `write-values`, `write-formulas`, `set-format`, `set-validation`, `set-conditional-formatting` | Записывает типизированные значения/формулы; задаёт формат, проверку или правило условного форматирования. | `action`, `workbook_id`, `sheet`, `address`, `payload?`, `patch?`, `spec?`, `mode?=add` | `RangeMutationResult` / `FormatMutationResult` / `ValidationRule` / `ConditionalFormatMutationResult` | Значения/формулы/формат/правила ячеек; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/F/F |
| `range_structural` | `copy`, `insert`, `unmerge` | Копирует диапазон; вставляет ячейки (`right`/`down`); снимает объединение. | `action`, `workbook_id`, `sheet`, `source?`, `destination_sheet?`, `destination?`, `address?`, `shift?`, `allow_merged?=false` | `StructuralMutationResult` | Данные получателя / сдвиг ячеек / разъединение; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/T/F |
| `range_delete` | (standalone) | Удаляет ячейки со сдвигом `left`/`up`. | `workbook_id`, `sheet`, `address`, `shift`, `allow_merged?=false` | `StructuralMutationResult` | Удаление ячеек; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/T/F |
| `range_merge` | (standalone) | Объединяет диапазон в одну ячейку. | `workbook_id`, `sheet`, `address`, `across?=false`, `allow_data_loss?=false` | `StructuralMutationResult` | Объединение; аудит | `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY`, `INVALID_RANGE`, `LISTOBJECT_INTERSECTION`, `PROTECTED_TARGET` | F/T/T |

`range_structural` **не идемпотентен** (H4): повторный `insert` дублирует, а `unmerge` отбрасывает
объединения. `range_write` не идемпотентен, поскольку `set-conditional-formatting(mode="add")`
добавляет правило (с перенумерацией приоритетов). Структурные операции, пересекающие Excel Table
(ListObject), отклоняются с `LISTOBJECT_INTERSECTION` и должны выполняться через инструменты
таблиц.

**Лимит ячеек (100 000).** Валидный диапазон, у которого `rows * cols` превышает `MAX_RANGE_CELLS`
(`100_000`), отклоняется **до выполнения** с явной ошибкой `RANGE_TOO_LARGE` (категория `validation`,
`isRetryable=false`) — никогда «молча» не урезается и никогда не сообщается как общий `INVALID_RANGE`
для корректного адреса. Агент должен разбить работу на меньшие диапазоны (страницами) через `offset` /
`limit`. Диапазон, обрезанный окном `offset`/`limit`, отмечается `truncated=true`, что отличается от
отклонения `RANGE_TOO_LARGE` по превышению лимита.

## Pivot

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `pivot_read` | `tables-list`, `table-get`, `cache-get`, `fields-list`, `field-get-filters` | Инвентарь сводных таблиц/кэшей; метаданные таблицы/кэша; список полей; состояние фильтров поля. | `action`, `workbook_id`, `pivot_table?`, `cache_index?`, `field?`, `background?=false` | `PivotInventory` / `PivotTableInfo` / `PivotCacheInfo` / `list[PivotFieldInfo]` / `PivotFieldFilters` (или `RunningJob`) | — | `COM_BUSY`, `AMBIGUOUS_FILTER` | T/F/T |
| `pivot_refresh` | `refresh-table`, `refresh-cache` | Обновляет одну сводную таблицу или один PivotCache. | `action`, `workbook_id`, `pivot_table?`, `cache_index?` | `PivotRefreshResult` | Данные сводной таблицы/кэша; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `REFRESH_FAILED`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/F/T |
| `pivot_field` | `field-set-layout`, `field-set-position`, `field-set-aggregation`, `field-set-properties`, `field-apply-filter`, `field-clear-filters` | Задаёт ориентацию/позицию/агрегацию/свойства поля; применяет/очищает фильтры. | `action`, `workbook_id`, `pivot_table`, `field?`, `orientation?`, `delete?=false`, `position?`, `function?`, `caption?`, `properties?`, `spec?`, `scope?=field` | `PivotFieldInfo` / `PivotFieldFilters` | Структура/данные сводной таблицы; аудит | `UNSUPPORTED_PIVOT_SOURCE`, `AMBIGUOUS_FILTER`, `INVALID_RANGE`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED`, `COM_BUSY` | F/T/T |

`pivot_read` — read-only; `pivot_refresh` и `pivot_field` — нет. `pivot_field` —
`destructive=true` **и** `idempotent=true` (двойная подсказка): `field-set-layout(delete=True)`
отбрасывает поле, а повторное применение той же раскладки/позиции/агрегации идемпотентно.
`cache-get` / `refresh-cache` требуют `cache_index`; действия над полями требуют
`pivot_table` + `field` (`field-clear-filters` также принимает table-scope). Создание/удаление
PivotTables/PivotCaches не поддерживается. `PivotFilters.Add` (метка/значение/дата/топ-низ) не
поддерживается на валидированной сборке: возвращается явный `unsupported`-результат, а не «тихая»
конвертация в ручной фильтр. Состояние поля страницы (report-filter) классифицируется по
`CurrentPage` / `AllItemsVisible`; `field-apply-filter` для поля страницы отклоняется до любой
мутации COM.

## Объекты

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `shape` | `list`, `get`, `set-on-action` | Все фигуры; одна фигура по имени; назначает макрос фигуре. | `action`, `workbook_id`, `name?`, `macro_name?` | `list[ShapeInfo]` / `ShapeInfo` | `OnAction`; аудит | `NOT_FOUND`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED` | F/F/T |
| `control` | `list`, `get`, `set-on-action` | Все Form Control/ActiveX; один элемент по имени; назначает макрос Form Control. | `action`, `workbook_id`, `name?`, `macro_name?` | `list[ControlEntry]` / `ControlEntry` | `OnAction`; аудит | `NOT_FOUND`, `WORKBOOK_READ_ONLY`, `SNAPSHOT_FAILED` | F/F/T |

`shape` и `control` — смешанные инструменты (list/get читают, set-on-action мутирует) и являются
`readOnly=false, destructive=false, idempotent=true` (все действия идемпотентны по матрице). Пустой
`macro_name` снимает назначение. ActiveX-элементы отклоняются: у них нет `OnAction`-семантики Form
Control. Отсутствующая фигура/элемент возвращают `NOT_FOUND` (F6); неоднозначное (повторяющееся)
имя — обычный `ExcelVbaMcpError` с семантикой `COM_BUSY` и `context.reason="ambiguous"`.

## Возможности / здоровье

| Инструмент | Действия | Назначение | Параметры | Возврат | Побочные эффекты | Ошибки | R/D/I |
|------------|----------|------------|-----------|---------|------------------|--------|-------|
| `xlvba_environment` | (standalone) | Сообщает режим доступа, политику, сборку и возможности сервера без обращения к COM. | — | `EnvironmentReport` | — | — | T/F/T |
| `server_health` | (standalone) | Read-only liveness-проба; никогда не обращается к Excel/COM. | — | `{"status", "service", "version"}` | — | — | T/F/T |

`xlvba_environment` регистрируется с **inline** `read_only`-аннотацией (не через общий набор
аннотаций); `server_health` также inline. Оба read-only и никогда не трогают COM.

## Подробно о сложных инструментах

### workbook `action="open"`

`workbook` `open` по умолчанию использует `mode="execute", read_only=False` — сессия для записи и
выполнения. Передайте `mode="inspect"` и/или `read_only=True` для сессии с отключёнными макросами
или только для чтения:

```json
{ "action": "open", "path": "C:/data/report.xlsm", "mode": "inspect", "read_only": true }
```

Ожидаемая форма ответа (`WorkbookInfo`):

```json
{
  "workbook_id": "1f0c…",
  "name": "report.xlsm",
  "format": "xlsm",
  "path": "<temp>/excel_vba_mcp/work/1f0c….xlsm",
  "source_path": "C:/data/report.xlsm",
  "read_only": true,
  "has_vba": true,
  "excel_version": "16.0.14332",
  "session_id": "9b2a…",
  "mode": "inspect"
}
```

`WorkbookInfo` несёт **два пути** и намеренно **не** содержит `original_name`:

- `path` — это **рабочая копия**, фактически открытая сервером (материализованный файл в каталоге
  `work/` сервера). Каждое чтение, запись и мутация нацелены на эту копию; оригинал никогда не
  изменяется.
- `source_path` — **исходный, allowlisted** путь, из которого была материализована рабочая копия,
  либо `null`, когда у книги нет источника на диске.
- `original_name` намеренно отсутствует (избыточен с `name`).

### workbook `action="close"` / `action="save"`

```json
{ "action": "close", "workbook_id": "1f0c…", "save_changes": false }
```

`close` без сохранения — read-only и разрешён; ответ — `SessionStatus` с `state="closed"`.
**По умолчанию отказ (D1):** при несохранённых изменениях `close` отказывает, если не задан
`discard_changes=true`; `save_changes=true` сохраняет при закрытии (мутация, `APPROVAL_DENIED` в
`restricted`); противоречивые флаги отклоняются. Явное сохранение:

```json
{ "action": "save", "workbook_id": "1f0c…" }
```

Требует approval (авторизовано в `full`; `APPROVAL_DENIED` в `restricted`) и возвращает обновлённый
`WorkbookInfo` с `dirty=false`. Изменения не сохраняются неявно никогда.

### range_read `action="read-values"` / range_write `action="write-values"`

Чтение ограниченного окна:

```json
{
  "action": "read-values",
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C10",
  "offset": 0,
  "limit": 10
}
```

Ответ — `RangeValues`:

```json
{
  "range": { "sheet": "Sheet1", "address": "A1:C10", "rows": 10, "cols": 3 },
  "shape": "block",
  "rows": 10,
  "cols": 3,
  "values": [
    [ { "kind": "string", "value": "Header" }, { "kind": "empty" }, { "kind": "empty" } ]
  ],
  "truncated": false,
  "notes": []
}
```

Значения типизированы: `string`, `number`, `boolean`, `date`, `error`, `empty`; потеря точности при
преобразовании в строку исключена. Запрос сверх лимита ячеек отклоняется, а не урезается;
`truncated=true` появляется, когда окно `offset`/`limit` обрезано.

Запись требует точного прямоугольника `payload`:

```json
{
  "action": "write-values",
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "B2:C3",
  "payload": {
    "rows": 2,
    "cols": 2,
    "values": [
      [ { "kind": "number", "value": 1 }, { "kind": "string", "value": "=1+1" } ],
      [ { "kind": "boolean", "value": true }, { "kind": "empty" } ]
    ]
  }
}
```

Строка, начинающаяся с `=`, сохраняется как **текст**, а не как формула (используйте
`range_write` `action="write-formulas"`). Неверный размер payload отклоняется с `INVALID_RANGE` без
изменений. Ответ — `RangeMutationResult` с `range_before`, `range_after`, `verified` и
`postcondition_values`.

### range_write `action="set-format"`

Разреженная (sparse) запись: применяются только явно заданные поля `patch`.

```json
{
  "action": "set-format",
  "workbook_id": "1f0c…",
  "sheet": "Sheet1",
  "address": "A1:C1",
  "patch": {
    "font_bold": true,
    "font_size": 11,
    "fill_color": 15773696,
    "number_format": "0.00",
    "alignment_horizontal": "center"
  }
}
```

Отсутствующие поля остаются без изменений; `null` сбрасывает значение в автоматическое только там,
где это допускает Excel.

### table `action="update-rows"` / table_delete_rows

Обновление допускает **явный индекс** или **однозначный `where`**:

```json
{
  "action": "update-rows",
  "workbook_id": "1f0c…",
  "name": "Sales",
  "updates": [
    { "index": 2, "values": { "Amount": 120 } },
    { "where": { "Region": "North" }, "values": { "Flag": "Y" } }
  ]
}
```

Если `where` не разрешается в ровно одну строку, инструмент возвращает `COM_BUSY` с сообщением о
неоднозначности и предпросмотром найденных строк — автоматического «угадывания» нет. Удаление строк
использует отдельный `table_delete_rows` с `indices` или `where`; отсутствие обоих — ошибка
валидации.

### snapshot_create / snapshot_restore

```json
{ "workbook_id": "1f0c…" }
```

Ответ — `SnapshotManifest`:

```json
{
  "snapshot_id": "snap-…",
  "created_at": "2026-11-21T18:41:56Z",
  "source_path": "C:/data/report.xlsm",
  "file_name": "report.xlsm",
  "stored_path": "<snapshots>/snap-…/report.xlsm",
  "sha256": "…",
  "size": 48213,
  "snapshot_method": "SaveCopyAs",
  "vba_exports": [
    {
      "component_name": "Module1",
      "type": 1,
      "path": "exports/Module1.bas",
      "export_file_sha256": "…"
    }
  ],
  "model_version": 2,
  "external_side_effects_not_undone": true
}
```

Верхнеуровневый `sha256` — хэш **файла** снапшота. `export_file_sha256` каждой записи `vba_exports`
— SHA-256 байтов **экспортированного файла компонента** (F7) — **не** валидный
`expected_source_hash` для инструментов мутации VBA, которым нужен нормализованный `source_hash` из
`vba_read` `action="read-component"`. `snapshot_list` возвращает те же манифесты; `snapshot_diff`
сравнивает полные манифесты и отдельно сообщает результат сравнения верхнеуровневых хэшей файлов.

Восстановление:

```json
{ "snapshot_id": "snap-…", "workbook_id": "1f0c…", "confirm": true }
```

`workbook_id` **обязателен** (текущая рабочая копия для отбрасывания). Текущая рабочая копия
закрывается **без сохранения**, затем открывается и валидируется новая копия из baseline; идентичность
восстановленной сессии перепривязывается к **оригиналу** (`source_path`), поэтому `list_workbooks` /
повторное открытие отражают оригинал, а не иммутабельный файл снапшота. Возвращается
`SnapshotRestoreResult`: восстановленный `WorkbookInfo` плюс `previous_close` (`"closed"` /
`"failed"` / `"not_requested"`) и, при ошибке закрытия, безопасный текст `previous_close_error`
(восстановление всё равно успешно, предыдущая копия остаётся открытой). Возвращённый `workbook.id` —
это **новый** активный id; используйте его для всех последующих вызовов, старый `workbook_id` закрыт.
Повторное восстановление уже активного baseline для того же оригинала отклоняется с
`reason="same_snapshot"`. Внешние эффекты макросов не откатываются. `confirm` обязателен.

### vba_replace_component_source

Перед заменой нужны два вызова `vba_read`: `list-components` (убедиться, что Standard/Class) и
`read-component` (получить `source_hash`). Затем:

```json
{
  "workbook_id": "1f0c…",
  "component": "Module1",
  "new_source": "Option Explicit\n\nPublic Sub Hello()\nEnd Sub",
  "expected_source_hash": "<source_hash из vba_read read-component>"
}
```

`expected_source_hash` — защита от устаревшей ревизии: сервер сравнивает его с текущим
нормализованным хэшем до и **внутри** STA-воркера. Несовпадение поднимает `REVISION_CONFLICT`.
Токен защиты — это **нормализованный `source_hash`** из `vba_read` `read-component`, а не хэш файла
экспорта снапшота. Document-модули и UserForms отклоняются.

### macro_run

```json
{ "workbook_id": "1f0c…", "macro_name": "Module1.Hello", "args": [] }
```

Порядок гейтов:

1. execute-режим + writable-сессия (`_require_execute_context`);
2. разрешение процедуры из инвентаря (`Module.Procedure`);
3. approval (ключ: canonical path + SHA-256 + действие + процедура) — авторизован в `full`,
   `APPROVAL_DENIED` в `restricted`;
4. `background_safe` для текущего хэша книги **и** `project_source_hash` (BG06);
5. preflight интерактивного UI (R13) — список токенов в [Сокращения](#сокращения);
6. один `Application.Run` на STA-воркере.

Возможные статусы `MacroRunResult`: `success`, `not_found`, `runtime_error`, `blocked`, `timeout`,
`needs_recovery`. При блокировке макрос **не запускается**. При дедлайне owned PID проверяется
только для чтения на наличие видимого модального окна (модальное → `MACRO_RUNTIME_ERROR`, нет →
`MACRO_TIMEOUT`); в обеих ветках завершается только принадлежащий автоматизации процесс Excel, и
сессия восстанавливается в процессе (`MacroRunResult.session_recovery` сообщает `"recovered"` /
`"needs_recovery"` / `null`). Восстановление **отбрасывает открытые книги**: прежний `workbook_id`
после этого недействителен, макрос никогда не переигрывается автоматически.

### macro `action="classify-background"`

```json
{ "action": "classify-background", "workbook_id": "1f0c…", "procedure": "Module1.Hello", "background_safe": true }
```

Необязательный `expected_project_source_hash` — защита от устаревшей ревизии (несовпадение →
`STALE_TARGET`). Классификация эвристической или интерактивной процедуры как `background_safe`
отклоняется с `INTERACTIVE_UI_NOT_SUPPORTED` (вето R13). Инструмент **не гейтится** и не является
границей безопасности — `macro_run` по-прежнему применяет approval-гейт и вето R13.

### test `action="run-suite"`

```json
{
  "action": "run-suite",
  "suite": {
    "suite": "smoke",
    "workbook_id": "1f0c…",
    "baseline": "create",
    "cases": [
      {
        "id": "reads-range",
        "steps": [
          { "action": "range_read", "sheet": "Sheet1", "range_address": "A1:B2", "limit": 2 }
        ],
        "assertions": [
          { "type": "range_structure", "sheet": "Sheet1", "address": "A1:B2", "rows": 2, "cols": 2 }
        ]
      }
    ]
  }
}
```

`action` и `type` берутся из **закрытых** словарей: `TestAction` и `AssertionType`; схема запрещает
лишние поля (`extra="forbid"`). Изоляция: один иммутабельный baseline-снапшот, затем **своя
одноразовая рабочая копия на каждый кейс**. Шаг `macro_run` может нести необязательное поле только
с данными `background_safe`, и раннер классифицирует процедуру через защищённый путь
`MacroService.classify_background` перед запуском. Компиляция перед макросом явная — нужен шаг
`vba_compile`. Прогон сохраняется (`runs/`), читается через `test` `action="get-run"`.

### vba_compile_diagnostics

```json
{ "workbook_id": "1f0c…" }
```

Fail-closed ответ:

```json
{
  "workbook_id": "1f0c…",
  "status": "indeterminate",
  "reason": "NON_ACTIVATING_COMPILE_UNAVAILABLE",
  "diagnostics": {
    "is_compile": false,
    "vbom_available": true,
    "protection": 0,
    "locked": false,
    "component_count": 3,
    "broken_references": []
  },
  "command_discovery": { "route_usable": false },
  "notes": [
    "NOT a compile: non-activating VBA compilation is unavailable on this build.",
    "status=indeterminate reason=NON_ACTIVATING_COMPILE_UNAVAILABLE"
  ]
}
```

Инструмент никогда не возвращает `compiled` или `compile_error`, не показывает и не активирует UI.
`vba_read` `action="get-compile-diagnostics"` возвращает те же данные из кэша без повторного сбора.

## Фоновые задачи

Четыре долгоживущих инструмента принимают `background=true`: `workbook` (`open`), `macro_run`,
`test` (`run-suite`) и `pivot_read` (`tables-list`). С `background=true` инструмент планирует
операцию в in-process реестре задач и возвращает promptly `RunningJob`-handle:

```json
{ "status": "running", "job_id": "job-…", "poll_after_s": 5 }
```

Клиент затем **pull/poll** `job_get(job_id)` не раньше `poll_after_s` (clamp `[1, 30]`, по
умолчанию `5`). Реестр **владеет результатом**, поэтому таймаут или дисконнект клиента никогда не
теряет его; завершённый результат остаётся доступным по `job_id` (и после перезапуска — из
персистированного `JobStore`).

Поля `JobStatus`: `job_id`, `status`, `started`, `cancel_requested`, `poll_after_s`, `result` (когда
терминально), `error_code`, `error_text`, `retry_after_s`, `phase`, `outcome`, `retryable` и
аддитивная история `events`.

`status` — одно из:

| Статус | Значение |
|--------|----------|
| `queued` | Запланирована, но ещё не запущена на STA-воркере. |
| `running` | Операция выполняется на единственном STA-воркере. |
| `succeeded` | Обёрнутая операция вернула значение (хранится в задаче). |
| `failed` | Обёрнутая операция подняла доменную или неожиданную ошибку. |
| `expired` | Операция превысила execution-дедлайн (`EXCEL_HUNG`); вызов оставлен выполняющимся на STA (никогда не отменяется). |
| `interrupted_unknown` | Персистированная нетерминальная задача, сверенная при старте, или задача, финализированная после успешного `excel_recover`. `outcome="unknown"`, `retryable=false`; никогда не завершена, никогда не переигрывается автоматически. |
| `unknown` | Зарезервировано; никогда не порождается реестром. |

`job_request_cancel` фиксирует только **намерение** отмены (`cancel_requested=true`); он не может
прервать непрерываемый COM-вызов, поэтому задача всё равно дорабатывает до терминального состояния.
Жизненный цикл и модель персистентности описаны в [Архитектура](04-architecture.md); взаимодействие
с `excel_recover` — в [Устранение неполадок](10-troubleshooting.md).

## Зависимости между инструментами

1. **`workbook` `action="open"` перед любым workbook-инструментом.** Все инструменты с `workbook_id`
   требуют открытой рабочей копии; иначе — `ExcelVbaMcpError` «unknown workbook_id».
2. **`snapshot_create` до любой guarded-мутации.** Без успешного снапшота мутация отклоняется с
   `SNAPSHOT_FAILED`.
3. **`vba_read` (`list-components` / `read-component`) до `vba_replace_component_source` и
   `vba_insert_code`.** Нужны имя компонента и актуальный `source_hash` для `expected_source_hash`.
4. **`vba_compile_diagnostics` до `vba_read` `get-compile-diagnostics`** — только чтобы «засеять»
   кэш; без первого вызова диагностика собирается заново (тот же результат).
5. **`macro` `list` до `macro_run`.** Позволяет узнать точное имя `Module.Procedure` и статус
   `background_safe` до запуска.
6. **`test` `get-run` после `test` `run-suite`.** `run_id` берётся из
   `TestRunResult.test_run_id`; инструмент только читает сохранённый прогон.

Дополнительно: `workbook` `close` (без сохранения) и `workbook` `save` несовместимы с последующими
вызовами по тому же `workbook_id` — после закрытия идентификатор больше не действителен.

## Мутирующие и потенциально опасные операции

| Инструмент | Что меняется | Почему гейтится | Как использовать безопасно |
|------------|--------------|-----------------|---------------------------|
| `workbook` `open` (writable/execute) | Создаёт рабочую копию и открывает её с макросами | Открытие по умолчанию; авторизовано в `full`, `APPROVAL_DENIED` в `restricted` | Использовать `mode="inspect", read_only=true` для read-only работы; `execute` применять только к проверенным копиям |
| `workbook` `close` (`discard_changes=true`) | Отбрасывает несохранённые изменения и закрывает | Разрушительно; по умолчанию отказ (D1) | Предпочесть `save_changes=true` или предварительный `snapshot_create`; передавать `discard_changes=true` только осознанно |
| `workbook` `save` | Записывает файл рабочей копии | Mutating | Комбинировать со снапшотом; не сохранять «на всякий случай» |
| `excel_save_as` | Записывает рабочую копию по целевому пути (own permission) | Standalone high-risk; overwrite policy | Проверить цель; использовать `overwrite_existing` / `overwrite_source` только после ревью |
| `snapshot_restore` | Отбрасывает текущую копию и открывает baseline | Mutating; требует `workbook_id` + `confirm=true` | Делать свежий `snapshot_create` сначала; отсутствующий `workbook_id` (`INVALID_RANGE`) или ложный `confirm` (`CONFIRMATION_REQUIRED`) отклоняется до отбрасывания; ошибка закрытия предыдущей копии сообщается (`previous_close="failed"`) и не ломает восстановление |
| `sheet_delete` | Удаляет лист с данными | Разрушительно, standalone | Проверять `sheet_read` `list`; полагаться на снапшот |
| `table_delete_rows` | Удаляет строки таблицы | Разрушительно, standalone (XOR-guard) | Использовать явные `indices`; `where` только при однозначности |
| `table_delete_column` | Удаляет колонку таблицы | Разрушительно, standalone (own permission) | Проверять `table_read` `get`; не передавать неоднозначный заголовок |
| `range_write` | Перезаписывает значения/формулы/формат/правила | Mutating | Проверять размер payload; делать снапшот |
| `range_structural` | Копирует/вставляет/разъединяет ячейки | Разрушительно; не идемпотентно | Передавать `shift` явно для `insert`; проверять пересечение с таблицами |
| `range_delete` | Удаляет ячейки со сдвигом | Разрушительно, standalone | Всегда задавать `shift` явно; проверять пересечение с таблицами |
| `range_merge` | Объединяет ячейки, отбрасывая данные вне якоря | Разрушительно (двойная подсказка) | Просматривать preconditions; передавать `allow_data_loss=true` только после ревью |
| `name_delete` | Удаляет определённое имя | Разрушительно, standalone | Проверять `name_read` и зависимости; полагаться на снапшот |
| `vba_replace_component_source` | Полностью заменяет код модуля | Разрушительно, standalone | Передавать точный `expected_source_hash`; проверять `vba_compile_diagnostics` после |
| `vba_insert_code` | Вставляет код в модуль | Разрушительно, standalone | Указывать точную `position`; не вставлять код «наугад» |
| `shape` `set-on-action` | Назначает макрос фигуре | Mutating | Проверять `shape` `list`; пустой `macro_name` снимает назначение |
| `control` `set-on-action` | Назначает макрос Form Control | Mutating | Работать только с Form Control; ActiveX не поддерживается |
| `macro_run` | Исполняет VBA-код | Executing; гейтится approval (restricted) + `background_safe` + R13 | Только после `macro` `list` / `macro` `classify-background` и preflight; не запускать макросы с внешними эффектами |
| `test` `run-suite` | Запускает декларативный набор (может диспетчеризовать `macro_run`) | Executing; destructive | Просматривать кейсы; шагу `macro_run` нужна классифицированная процедура |
| `excel_recover` | Завершает owned-процесс Excel и отбрасывает рабочие копии | Recovery; allow-by-default (оба режима), аудируется, pid-scoped, fail-closed; требует `confirm=true` | Использовать только когда `excel_health` показывает `needs_recovery` / `running/deadline_exceeded`; оригиналы никогда не изменяются, но открытые рабочие копии теряются |

## Read-only ресурсы

| URI template | Name | MIME | Содержимое |
|--------------|------|------|------------|
| `excel://workbooks/{workbook_id}/model` | `workbook_model` | `application/json` | Агрегированная модель книги Stage 1 (листы, таблицы, имена, фигуры, элементы, опционально VBA) |
| `excel://workbooks/{workbook_id}/vba/components` | `workbook_vba_components` | `application/json` | Инвентарь VBA-компонентов + `vba_access` |
| `excel://workbooks/{workbook_id}/vba/components/{component_name}` | `workbook_vba_component_source` | `text/plain` | Исходный текст одного VBA-компонента |
| `excel://test-runs/{run_id}` | `test_run_summary` | `application/json` | Сохранённый декларативный тестовый прогон |

Ресурсы доступны только для чтения: они не мутируют книгу и не выполняют VBA.

## Известные ограничения инструментов

1. **Компиляция всегда `indeterminate`.** `vba_compile_diagnostics` не может подтвердить компиляцию
   на валидированной сборке; `is_compile=false`.
2. **Макросы требуют классификации `background_safe`.** `BackgroundSafetyRegistry` стартует пустым
   (fail-closed); классифицируйте процедуру через `macro` `classify-background`.
   Неклассифицированная процедура возвращает `blocked` (`BACKGROUND_SAFETY_NOT_VERIFIED`);
   интерактивная блокируется вето R13 (`INTERACTIVE_UI_NOT_SUPPORTED`). Токен-скан R13 — эвристический
   запрет, а не доказательство.
3. **Нет grant-инструмента для approval.** Управление approval/allowlist через MCP невозможно. В
   режиме `full` (по умолчанию) мутации и execute-режим авторизованы без гранта; в `restricted` они
   отклоняются с `APPROVAL_DENIED`. Следовательно, в режиме `restricted` нет in-band пути для выдачи
   approval или управления allowlist, поэтому execute-`workbook` `open` по умолчанию отклоняется и
   пригодны только сессии `inspect`+`read_only`; решение — работать в режиме `full` или заранее
   авторизовать через библиотеку.
4. **Нет создания/удаления PivotTables и PivotCaches.** Только управление существующими объектами;
   внешние/OLAP/Data-Model источники отклоняются.
5. **Пересечение диапазона с ListObject.** Структурные операции диапазонов над таблицами
   отклоняются с `LISTOBJECT_INTERSECTION`; нужно использовать инструменты таблиц.
6. **Ограниченные чтения.** `range_read` (`read-values` / `read-formulas`) и `table_read`
   (`read-rows`) возвращают ограниченные окна; большие диапазоны нужно читать порциями через
   `offset`/`limit`.
7. **Нет диаграмм, сводных диаграмм и слайсеров.** Инструменты для них не предоставляются.
8. **Нет создания книг.** `workbook` `open` открывает только существующие файлы; создать новую книгу
   через MCP нельзя.
9. **Нет полного пересчёта и cross-workbook копирования.** `Application.Calculate` не предоставляется;
   `range_structural` `copy` работает только внутри одной книги и без буфера обмена.
10. **Нет пагинации списков.** `table_read` `list` и `pivot_read` `tables-list` возвращают полный
    ограниченный список без постраничного вывода.
11. **Потерянный owned-Excel обнаруживается как `needs_recovery`; восстановление agent-directed.**
    `excel_health` читает bookkeeping воркера/сессии без обращения к COM и дополнительно выполняет
    read-only liveness-пробу owned-pid: если принадлежащий скрытый `EXCEL.EXE` умер, а bookkeeping
    всё ещё говорит `running`/`accepting`, `excel_health` сообщает `needs_recovery` вместо ложного
    `ok`. Зависание/таймаут `macro_run` пересоздаёт экземпляр автоматически (F2b). Для любого
    другого состояния `needs_recovery` / `running/deadline_exceeded` агент вызывает `excel_recover`
    (agent-directed, fail-closed, originals-safe); сервер никогда не завершает живой pid по дедлайну.
    Если восстановление отклонено, fallback — перезапуск сервера.
12. **У `macro_run` нет предусловия снапшота (задокументированное отклонение от спеки).** В отличие
    от контент-мутирующих инструментов, `macro_run` не требует предварительного `snapshot_create`.
    Рекомендованный процесс всё равно делает снапшот первым, чтобы внешние эффекты макроса имели
    baseline для отката.
13. **`excel_recover` — единственный recovery-инструмент; инструмента `excel_recover_session` нет.**
    `SessionService.recover_session`/`close_session` существуют внутри, но не выставлены как
    MCP-инструменты. Путь зависания/таймаута `macro_run` также восстанавливает автоматически (F2b) и
    сообщает об этом в `MacroRunResult.session_recovery`. Прерванная операция сообщается как
    `prior_job="unknown"` и никогда не переигрывается автоматически.
14. **`macro_policy.classify_operation` неавторитетен (мёртвая/иллюстративная таблица).** Функция
    `classify_operation` в `security/macro_policy.py` сохранена для иллюстрации и **не** подключена к
    пути рантайм-энфорсмента; живой класс операции каждого инструмента определён по-сервисному. Её
    таблица предшествует текущему каталогу из 41 инструмента и не должна читаться как авторитетная
    классификация инструментов — используйте этот справочник или код сервисов.