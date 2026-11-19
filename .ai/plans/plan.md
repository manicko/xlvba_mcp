# Excel VBA MCP Server — Implementation Execution Plan

## 1. Scope and rules

**Implementation objective:** дать MCP-совместимому AI-агенту контролируемый локальный доступ к настольному Microsoft Excel, чтобы агент мог исследовать книгу, читать и изменять VBA-код, компилировать VBA-проект средствами VBE, запускать разрешённые макросы, создавать воспроизводимые тестовые копии, проверять результаты и восстанавливать исходное состояние.

**Specification source:** `excel_vba_mcp_technical_spec.docx` (далее — Спецификация). План не меняет и не расширяет заданную архитектуру.

**В объёме:** этап 1 и обязательная инфраструктура Workbook / Excel Session. Этап 2 в задачи реализации не декомпозируется.

**Out of scope for Stage 1:** полный Range/Cells API; PivotTables; Charts и Slicers; Connections и Power Query; Data Model / Power Pivot; дополнительные объекты Excel; полноценные UI-click тесты ActiveX. Для проверки выходов разрешены только явно указанные read-only assertions.

**Безопасное исполнение:** оригиналы в `projects/` не изменять; мутации выполнять в разрешённых рабочих копиях; перед изменением VBA/структуры книги и перед каждым тестовым сценарием должен существовать baseline snapshot.

**Product Owner questions:** не требуются. В Спецификации нет неразрешённых бизнес-правил или неопределённостей пользовательского поведения. Оставшиеся вопросы являются техническими и закрываются PoC/research gates.

## 2. Agent reviews and research status

В доступном наборе инструментов этой сессии нет отдельного механизма запуска независимых Auditor / Researcher / Planner / Validator subagents. Эти агенты не запускались; их ревью не имитируется. Для технических развилок проверены первичная документация Microsoft, официальный MCP Python SDK и документация pywinauto. В задачах явно обозначено, где перед реализацией требуется независимое ревью.

Порядок gate для задач, где требуются все роли: **Researcher → Auditor → Planner → Validator → implementation**. Researcher сравнивает подходы в рамках уже заданной архитектуры; Auditor проверяет предпосылки и границы модулей; Planner уточняет конкретную последовательность и тесты; Validator независимо проверяет решение и критерии приёмки. Если остаётся неопределённость, она фиксируется как открытый PoC/decision gate; конфликтующие варианты не смешиваются в одной задаче.

High-risk задачи требуют всех четырёх ролей. Для задач средней сложности используются только указанные роли. Для простых задач отдельные агенты не требуются.

## 3. Risk register and unresolved options

| ID | Risk / uncertainty | Compatible options to validate | Execution rule within the specification |
|---|---|---|---|
| R01 | COM STA and message dispatch | Dedicated STA worker with queued calls and message pumping; behavior under `RPC_E_CALL_REJECTED` and modal dialogs | Preserve `com_worker.py` and a single STA worker. Prove its lifecycle on target Windows/Excel; COM proxies never leave the worker. |
| R02 | VBIDE access and component export | `VBProject / VBComponents / CodeModule`; VBIDE Export; textual parsing as marked heuristic fallback | VBIDE is source of truth. Validate UserForm `.frm` + `.frx`; preserve Document Modules. |
| R03 | VBE compile command discovery | Localized VBE CommandBars lookup; checked control ID as fallback | Do not rely on a fixed identifier alone. Validate on target Office build and UI language. Never return `compiled` without evidence. |
| R04 | Compile error UI extraction | `pywinauto` UIA backend; direct Windows UI Automation only if the first route is insufficient | Start with the specified pywinauto/UIA path. Do not introduce an alternate mechanism without PoC evidence. No coordinate/SendKeys-first solution. |
| R05 | Macro runtime errors | COM return/exception; COM plus UIA modal-window observation | Keep `Application.Run`. Verify actual error evidence; do not promise the original VBA runtime description when Excel returns a generic COM error. |
| R06 | Snapshot consistency | `Workbook.SaveCopyAs` for an open workbook; file copy for a closed workbook | Follow the two branches in the specification. Each test case gets its own disposable copy. Snapshot is not a sandbox for external side effects. |
| R07 | ListObject mutations | `ListRows` / `ListColumns` plus postcondition checks | Mutate the ListObject, not adjacent cells. Ambiguous row updates/deletes must not execute. |
| R08 | Non-range Names | Attempt `RefersToRange` in an isolated error branch; retain `RefersTo` for non-range Names | Do not assume every Name points to a Range; do not evaluate formulas with `eval`. |
| R09 | ActiveX event mapping | Enumerate `OLEObjects`; find probable `ControlName_Event` procedure in the corresponding Document Module | Label inferred mappings as probable; no automatic ActiveX click testing in MVP. |
| R10 | Untrusted VBA and external effects | Approval or allowlist by workbook path/hash + procedure; interactive user session | Snapshot does not roll back databases, email, network/file side effects. No Windows Service or unattended Office automation. |

## 4. Execution DAG

Excel COM calls on one target desktop are serialized through the single STA worker. “Parallel” below means independent code/review preparation; it does not mean concurrent Excel automation against the same instance.

| Group | Tasks | Gate / dependency | Parallel execution |
|---|---|---|---|
| G0 Bootstrap | `task_001` | Reproducible environment and MCP STDIO skeleton | Base project only |
| G1 Risk-reduction PoC | `task_002`–`task_006` | Go/No-Go for Excel COM, VBIDE, VBE compile/UIA, Application.Run and SaveCopyAs | Prepare fixtures/reviews in parallel; execute Office UI tests serially |
| G2 Contracts and policies | `task_007`–`task_009` | DTOs, path policy and approval/audit ready | Independent tasks after bootstrap |
| G3 Excel runtime | `task_010` → `task_011` → `task_012` → `task_013` | STA worker, worker verification, session service, session verification | Sequential integration chain |
| G4 Snapshot foundation | `task_014` → `task_015` | Snapshot store/service and baseline verification | After session service; read-only VBA work can proceed separately |
| G5 VBA implementation | `task_016` → `task_017`; `task_018` → `task_019` | Inventory before source mutation; snapshot before mutation | Read/export and object read-only work can be developed in parallel |
| G6 Compile and macro | `task_020` → `task_021`; `task_022` → `task_023` | Compile PoC and VBA inventory before compile service; compile gate before test-run macro execution | Compile/macro service work follows their own prerequisites |
| G7 Workbook objects | `task_024`–`task_031` | Session, DTO, security; snapshot for mutating operations; VBA inventory for control mapping | Sheets, Tables, Names, Controls are separate branches; Excel integration runs remain serialized |
| G8 Unified model | `task_032` → `task_033` | All Stage 1 inventory services complete | Aggregation after underlying services to avoid temporary duplicate contracts |
| G9 Test harness | `task_034` → `task_035` → `task_036` | Recovery/error and artifact foundation → declarative runner → real end-to-end verification | Schema/review can start earlier; integration waits for dependencies |
| G10 Release | `task_037` | All A01–A15 evidence available | Final regression only after feature freeze |

**Critical path:** `task_001` → `task_002` → `task_010` → `task_011` → `task_012` → `task_014` → `task_016` → `task_018` → `task_020` → `task_022` → `task_032` → `task_034` → `task_035` → `task_036` → `task_037`.

`task_003`–`task_006` are additional early Go/No-Go gates. A failed gate blocks its dependent implementation; it is not permission to silently substitute a different architecture or to report an unverified capability as complete.

---

# 5. Implementation task specifications

## task_001 — Bootstrap the local MCP project

```yaml
id: task_001
title: Bootstrap the local Excel VBA MCP project
priority: high
depends_on: []
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "4. Технологический стек; 5.1. Структура каталогов; 6. Контракт MCP"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["Таблица 5 — Предлагаемый стек", "Таблица 7 — Структура каталогов", "Таблица 8 — Ответственность ключевых компонентов"]
description: >
  Создать воспроизводимую структуру Python-проекта и минимальный MCP STDIO процесс.
  Использовать каталоги и границы модулей из Спецификации; не добавлять альтернативный runtime/transport.
goals:
  - создать pyproject.toml и uv.lock с фиксируемыми зависимостями
  - использовать Python 3.14 x64 как целевой runtime (нижняя граница проекта — 3.11;
    изначальная цель 3.12 ослаблена до 3.14 по решению от 2026-11-19, см. execution log)
  - создать __main__.py, server.py и config.py с минимальным запуском STDIO
  - исключить обычные логи из stdout
  - создать projects, work, snapshots, artifacts и tests
extra_context: |
  Версии пакетов не следует оставлять плавающими в production. После проверки конкретного
  окружения зафиксировать точные версии в uv.lock. Стек: Windows 11 x64, Python 3.14 x64
  (нижняя граница проекта 3.11; runtime зафиксирован в .venv), uv, официальный MCP Python SDK v2,
  pywin32, pywinauto, Pydantic v2, pytest.
files:
  - path: pyproject.toml
    targets: [{type: project_config, name: dependencies and entry points}]
  - path: uv.lock
    targets: [{type: dependency_lock, name: pinned environment}]
  - path: src/excel_vba_mcp/__main__.py
    targets: [{type: module, name: process entry point}]
  - path: src/excel_vba_mcp/server.py
    targets: [{type: module, name: MCPServer and STDIO lifecycle}]
  - path: src/excel_vba_mcp/config.py
    targets: [{type: module, name: typed configuration}]
  - path: README.md
    targets: [{type: section, name: local installation and start}]
changes:
  - action: add_module
    description: Создать package skeleton и минимальный MCP server lifecycle.
  - action: configure_dependencies
    description: Зафиксировать проверенные версии; не переносить старый FastMCP v1 пример без адаптации.
acceptance_criteria:
  - uv sync --locked работает в чистой Windows среде
  - MCP server запускается через STDIO и обнаруживается MCP Inspector
  - stdout содержит только сообщения протокола
  - окружение воспроизводится из uv.lock
tests_to_run: ["uv sync --locked", "package import/startup smoke test", "MCP Inspector discovery"]
agents: []
verification_mode: inline
```

## task_002 — Prove isolated Excel COM and STA behavior

```yaml
id: task_002
title: Prove isolated Excel COM and STA behavior on the target machine
priority: critical
depends_on: [task_001]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "5.3. COM и параллелизм; 7.1. Workbook / Excel Session; 13.1. Обязательные Proof-of-Concept"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["Таблица 20 — Стадия 0. Discovery / PoC"]
description: >
  Выполнить ограниченный PoC на целевой Windows/Excel конфигурации: DispatchEx, COM STA,
  сериализованный доступ, открытие/закрытие тестовой книги и независимость от пользовательской сессии.
goals:
  - зафиксировать Windows build, Office build, bitness и язык VBE
  - подтвердить отдельный Excel instance через DispatchEx
  - проверить CoInitializeEx(COINIT_APARTMENTTHREADED) и корректное завершение COM
  - проверить STA message pumping и поведение при COM/UI событиях
  - зафиксировать Go/No-Go и HRESULT
extra_context: |
  Все COM proxies создаются, читаются и освобождаются в одном STA worker. Инициализация:
  pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED); создание Excel через
  win32com.client.DispatchEx("Excel.Application"); завершение с закрытием owned workbooks,
  Quit и освобождением ссылок в finally. COM proxies не передаются в MCP слой.
files:
  - path: tests/integration_excel/test_excel_instance_poc.py
    targets:
      - {type: test_module, name: Excel instance and STA PoC}
      - {type: test_case, name: test_dispatch_ex_creates_isolated_excel_instance}
      - {type: test_case, name: test_excel_workbook_open_close_lifecycle}
changes:
  - action: add_poc
    description: Создать временный PoC harness в существующей integration_excel области, не production abstraction.
  - action: record_compatibility_result
    description: Зафиксировать build/bitness/language, COM behavior, failures and Go/No-Go.
acceptance_criteria:
  - отдельный Excel instance создан и идентифицирован
  - существующая пользовательская Excel-сессия не затронута
  - open/close/Quit не оставляют orphaned owned Excel process
  - COM proxies остаются в создавшем их STA thread
  - поведение message pump и ограничения задокументированы
tests_to_run: ["test_excel_instance_poc.py on target Windows/Excel"]
agents:
  - researcher: required_before_poc; compare compatible STA/message-pump practices
  - auditor: required; inspect COM lifetime and isolation assumptions
  - planner: required; define bounded PoC and failure capture
  - validator: required_after_research_and_planner; verify proof of isolation
verification_mode: dedicated_gate
```

## task_003 — Prove VBIDE access and component export

```yaml
id: task_003
title: Prove VBIDE access and component export on the target Excel build
priority: critical
depends_on: [task_002]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.2. VBA Project; 12.1. Test fixtures; 13.1. Обязательные Proof-of-Concept"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["CleanProject.xlsm fixture requirements"]
description: >
  Проверить AccessVBOM и реальный round-trip чтения/экспорта VBProject components через VBIDE.
  Trust Center автоматически не изменять.
goals:
  - распознавать отключённый AccessVBOM как ACCESS_VBOM_DISABLED
  - перечислить Standard, Class, UserForm и Document Modules
  - прочитать CodeModule и экспортировать .bas/.cls
  - экспортировать UserForm как пару .frm + .frx
  - сохранить Document Modules как компоненты книги
  - проверить References и IsBroken
extra_context: |
  Источник — Workbook.VBProject → VBComponents → CodeModule и References. AccessVBOM включается
  вручную пользователем; сервер только диагностирует недоступность. UserForms экспортировать
  с .frx; Document Modules не удалять и не пересоздавать.
files:
  - path: tests/fixtures/CleanProject.xlsm
    targets: [{type: fixture, name: CleanProject}]
  - path: tests/integration_excel/test_vbide_access_poc.py
    targets:
      - {type: test_module, name: VBIDE access and component export PoC}
      - {type: test_case, name: test_vbproject_access_diagnostic}
      - {type: test_case, name: test_userform_frm_frx_export_pair}
changes:
  - action: add_poc
    description: Проверить доступ, mapping типов, CodeModule read и export без production VBA service.
acceptance_criteria:
  - выключенный AccessVBOM обнаруживается без изменения реестра
  - типы всех поддерживаемых компонентов различаются корректно
  - экспорт UserForm содержит .frm и .frx
  - exported source читается обратно с согласованным hash
  - locked project возвращает VBA_PROJECT_LOCKED без обхода защиты
tests_to_run: ["test_vbide_access_poc.py on CleanProject.xlsm and locked fixture when available"]
agents:
  - researcher: required_before_poc; investigate supported VBIDE access/export limits
  - auditor: required; inspect Document Module and designer-resource risks
  - planner: required; define fixture coverage and evidence
  - validator: required_after_research_and_planner; review round-trip completeness
verification_mode: dedicated_gate
```

## task_004 — Prove VBE compilation and UIA diagnostics

```yaml
id: task_004
title: Prove VBE Compile command and UI Automation diagnostics
priority: critical
depends_on: [task_002, task_003]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "8. Компиляция VBA; 11.1. Ограничения; 12.1. Test fixtures; 13.1. Обязательные Proof-of-Concept"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["CompileErrorProject.xlsm fixture requirements"]
description: >
  Проверить на целевом Office build вызов VBE Debug → Compile и извлечение реального compile
  error dialog, выбранного контекста и screenshot средствами UI Automation.
goals:
  - проверить поиск Compile command через VBE CommandBars
  - проверить локализованную подпись/структуру и control ID fallback только после проверки
  - сравнить pywinauto UIA с прямым Windows UI Automation только при необходимости
  - отличать compiled, compile_error, already_compiled и indeterminate
  - подтвердить, что следующий COM вызов блокируется до разрешения modal dialog
  - зафиксировать UI language, selectors, dialog text и screenshot
extra_context: |
  В Excel Object Model нет публичного Workbook.CompileProject() / Application.CompileProject().
  Компиляция выполняется командой VBE Debug → Compile. При подтверждённом запуске и отсутствии
  ошибки — compiled; если проект не изменился и команда уже выключена — already_compiled;
  если команда/UIA/диалог не подтверждены — indeterminate. Нельзя выдавать compiled только
  из-за отсутствия COM exception. UIA — основной интерфейс; SendKeys/координаты не основной путь.
files:
  - path: tests/fixtures/CompileErrorProject.xlsm
    targets: [{type: fixture, name: CompileErrorProject}]
  - path: tests/integration_excel/test_vbe_compile_poc.py
    targets:
      - {type: test_module, name: VBE compile and UIA PoC}
      - {type: test_case, name: test_compile_success_fixture}
      - {type: test_case, name: test_compile_error_dialog_capture}
changes:
  - action: add_poc
    description: Сравнить допустимые варианты поиска команды и UIA diagnostics на целевом Office build.
  - action: record_open_choice
    description: Зафиксировать работавший вариант, fallback, build и непроверенные версии/локали.
acceptance_criteria:
  - success fixture returns compiled or truthful indeterminate
  - error fixture never returns false compiled
  - dialog text and screenshot are captured
  - selected token/context is captured when the UIA provider exposes it
  - no COM operation runs while compile dialog is unresolved
tests_to_run: ["test_vbe_compile_poc.py against success and compile-error fixtures"]
agents:
  - researcher: required_before_poc; compare VBE command lookup and UIA paths within specified adapter
  - auditor: required; inspect false-success and modal deadlock risks
  - planner: required; prepare success/error fixture matrix
  - validator: required_after_research_and_planner; challenge evidence sufficiency
verification_mode: dedicated_gate
```

## task_005 — Prove Application.Run behavior

```yaml
id: task_005
title: Prove Application.Run execution and error reporting
priority: high
depends_on: [task_002, task_003]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.3. Macro Execution; 11.1. Ограничения; 13.1. Обязательные Proof-of-Concept"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Проверить Application.Run на fixture procedures: позиционные аргументы, сериализуемые типы,
  Variant return value, missing macro, runtime error и возможный modal dialog.
goals:
  - проверить qualified workbook/procedure name
  - проверить scalar/date/array arguments и Variant return
  - проверить missing procedure и unhandled runtime error
  - зафиксировать COM и UIA evidence при modal error
extra_context: |
  Application.Run использует позиционные аргументы, максимум 30; принимать string, number,
  boolean, null, ISO date и поддерживаемые одномерные/двумерные массивы. Не принимать Python
  objects или COM references. Application.Run не является компилятором.
files:
  - path: tests/fixtures/CleanProject.xlsm
    targets: [{type: fixture, name: CleanProject test macros}]
  - path: tests/integration_excel/test_macro_run_poc.py
    targets:
      - {type: test_module, name: Application.Run PoC}
      - {type: test_case, name: test_macro_arguments_and_return_value}
      - {type: test_case, name: test_macro_runtime_error_behavior}
changes:
  - action: add_poc
    description: Запускать только тестовые fixture macros; записывать фактическое поведение Excel/COM/UIA.
acceptance_criteria:
  - successful macro returns serializable result and duration
  - unsupported argument types are rejected
  - missing macro/runtime error are distinguished when Excel exposes sufficient evidence
  - modal error is detected and session remains diagnosable
  - no claim is made that snapshot reverses external side effects
tests_to_run: ["test_macro_run_poc.py against CleanProject.xlsm"]
agents:
  - researcher: required_before_poc; inspect Application.Run argument and error behavior
  - auditor: required; assess modal and side-effect risks
  - planner: required; define safe fixture procedures and evidence
  - validator: required_after_research_and_planner; review classification and safety
verification_mode: dedicated_gate
```

## task_006 — Prove SaveCopyAs and restore behavior

```yaml
id: task_006
title: Prove workbook snapshot and restore behavior
priority: high
depends_on: [task_002]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "9.1. Snapshot model; 9.4. Restore, side effects and limits; 13.1. Обязательные Proof-of-Concept"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Проверить SaveCopyAs для открытой macro-enabled книги и файловое копирование для закрытой;
  подтвердить открытие копии и идентичность baseline по SHA-256.
goals:
  - проверить SaveCopyAs на .xlsm
  - открыть snapshot как отдельную disposable copy
  - закрыть изменённую copy без сохранения и повторно открыть baseline
  - проверить manifest и SHA-256
extra_context: |
  Для открытой книги использовать Workbook.SaveCopyAs, для закрытой — файловое копирование.
  Snapshot содержит файл, manifest.json, SHA-256, timestamp, source path, Excel version,
  workbook model и доступные VBA exports. Snapshot не отменяет внешние побочные эффекты.
files:
  - path: tests/integration_excel/test_snapshot_poc.py
    targets:
      - {type: test_module, name: snapshot and restore PoC}
      - {type: test_case, name: test_save_copy_as_reopen}
      - {type: test_case, name: test_restore_baseline_hash}
changes:
  - action: add_poc
    description: Подтвердить snapshot semantics до разработки snapshot_service.py.
acceptance_criteria:
  - SaveCopyAs copy открывается как отдельная workbook
  - baseline hash не меняется после мутации working copy
  - restore возвращает workbook file к baseline
  - вывод теста явно указывает, что external side effects не отменяются
tests_to_run: ["test_snapshot_poc.py on macro-enabled fixture"]
agents:
  - researcher: required_before_poc; verify supported snapshot-copy semantics
  - auditor: required; inspect file consistency and restore failure paths
  - planner: required; define mutation and restore evidence
  - validator: required_after_research_and_planner; review baseline identity criteria
verification_mode: dedicated_gate
```

## task_007 — Define Phase 1 DTOs and result contracts

```yaml
id: task_007
title: Define serializable Phase 1 domain models and statuses
priority: high
depends_on: [task_001]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "5.2. Ответственность ключевых компонентов; 6. Контракт MCP; 8.2. Диагностический формат; 9.2–9.3. Test suite and assertions; 11. Надёжность"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Описать Phase 1 DTO, статусы и типизированные ошибки в существующем domain слое.
  Контракты отражают только требования Спецификации.
goals:
  - определить DTO для workbook/session, VBA component/procedure/reference, compile, macro, snapshot, test run, sheet, table, name и control
  - определить error/status codes из Спецификации
  - определить bounded declarative test_spec без произвольного Python/shell кода
  - обеспечить JSON serialization и MCP structured outputs
extra_context: |
  COM objects не передаются за пределы STA worker; между слоями передаются Python DTO:
  strings, numbers, booleans, arrays, ISO dates, dict/list и status values.
files:
  - path: src/excel_vba_mcp/domain/models.py
    targets: [{type: module, name: Phase 1 serializable DTOs}]
  - path: src/excel_vba_mcp/domain/errors.py
    targets: [{type: module, name: typed errors}]
  - path: src/excel_vba_mcp/domain/statuses.py
    targets: [{type: module, name: operation/test statuses}]
  - path: src/excel_vba_mcp/domain/test_spec.py
    targets: [{type: module, name: declarative test suite schema}]
changes:
  - action: add_models
    description: Определить Pydantic v2 schemas, соответствующие Phase 1 contracts.
  - action: add_status_mapping
    description: Сопоставить документированные ошибки с устойчивыми error codes без потери HRESULT.
acceptance_criteria:
  - DTO сериализуются в JSON и не содержат COM objects
  - статусы включают все указанные в Спецификации error conditions
  - test spec принимает только документированные bounded actions/assertions
  - Phase 2 contracts не добавлены
tests_to_run: ["unit tests for validation and JSON serialization"]
agents: []
verification_mode: inline
```

## task_008 — Implement canonical path policy

```yaml
id: task_008
title: Implement canonical path and workbook allowlist policy
priority: critical
depends_on: [task_001, task_007]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.1. Workbook / Excel Session; 10. Безопасность и модель доверия"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать централизованную проверку путей к книгам и артефактам в allowlisted roots
  projects/, work/ и snapshots/.
goals:
  - нормализовать canonical path до операций с файлами/Excel
  - блокировать path traversal, symlink/reparse-point escape и UNC по умолчанию
  - проверять разрешённые extensions и тип операции
  - возвращать PATH_NOT_ALLOWED с безопасной диагностикой
extra_context: |
  Открывать только файлы с расширениями и путями из allowlist; нормализовать путь;
  запрещать выход за разрешённую директорию, учитывая symlink/reparse point.
files:
  - path: src/excel_vba_mcp/security/path_policy.py
    targets:
      - {type: class, name: PathPolicy}
      - {type: method, name: validate_path}
changes:
  - action: add_policy
    description: Реализовать canonical path resolution и operation-specific allowlist checks.
acceptance_criteria:
  - paths outside configured roots are rejected
  - parent traversal and symlink/reparse escapes are rejected
  - UNC paths are denied by default
  - permitted paths return canonical value suitable for workbook locking
  - arbitrary filesystem paths are never sent directly to Excel COM
tests_to_run: ["unit tests for normal, traversal, symlink/reparse and UNC paths"]
agents:
  - auditor: recommended; review filesystem boundary and reparse-point cases
  - validator: recommended; review path-policy regression tests
verification_mode: inline
```

## task_009 — Implement approval policy and audit trail

```yaml
id: task_009
title: Implement approval policy, macro policy and operation audit logging
priority: critical
depends_on: [task_007, task_008]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.1. Workbook / Excel Session; 7.2. VBA Project; 7.3. Macro Execution; 10. Безопасность и модель доверия"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["Таблица 17 — Безопасность и модель доверия"]
description: >
  Реализовать policy gate и audit trail для операций, меняющих VBA/книгу или исполняющих макрос.
  Сохранить inspect/execute separation и не изменять Trust Center.
goals:
  - классифицировать операции как read-only, mutation или execution
  - требовать approval/allowlist для VBA write, структурных изменений и macro run
  - поддержать allowlist по workbook path/hash + action/procedure
  - логировать workbook hash, tool, редактированные параметры, approval, snapshot, время, результат и HRESULT
  - диагностировать AccessVBOM и macro policy без изменения настроек Office
extra_context: |
  Чтение модели/кода — read-only. Запись VBA, изменение таблиц/Names/Sheets, удаление объектов
  и запуск макроса требуют локального approval или allowlist по workbook hash + action.
  Не включать все макросы; не устанавливать AutomationSecurity=Low как общее правило; не менять
  registry, AccessVBOM, MacroSettings или Trusted Locations автоматически; не снимать MOTW.
  Не обходить пароль VBA проекта. Не сохранять credentials в логах.
files:
  - path: src/excel_vba_mcp/security/macro_policy.py
    targets: [{type: class, name: MacroPolicy}]
  - path: src/excel_vba_mcp/security/approval.py
    targets: [{type: class, name: ApprovalPolicy}]
  - path: src/excel_vba_mcp/security/audit.py
    targets: [{type: class, name: AuditLogger}]
changes:
  - action: add_policy
    description: Реализовать typed decisions для read/mutate/execute, привязанные к workbook identity и requested action.
  - action: add_audit_events
    description: Писать JSON audit events без секретов и без использования MCP stdout.
acceptance_criteria:
  - unapproved mutation/execution is denied
  - approved action is auditable with workbook identity and result
  - inspect mode does not enable macros
  - Trust Center/registry are never modified
  - VBA project protection is never bypassed
tests_to_run: ["unit tests for allow/deny, audit redaction and workbook hash matching"]
agents:
  - researcher: required; verify Office macro trust boundaries compatible with the specification
  - auditor: required; review threat model and log redaction
  - planner: required; map every operation to policy class
  - validator: required_after_research_and_planner; challenge deny-by-default and approval bypass tests
verification_mode: dedicated_gate
```

## task_010 — Implement the dedicated Excel COM STA worker

```yaml
id: task_010
title: Implement the serialized Excel COM STA worker
priority: critical
depends_on: [task_002, task_007]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "3. Ключевые архитектурные решения; 5.2–5.3. Components and COM concurrency"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать production COM worker по подтверждённому PoC: один выделенный STA thread,
  сериализованная очередь запросов и контролируемый lifecycle COM.
goals:
  - инициализировать COM в worker thread
  - выполнять все Excel COM operations только в worker
  - сериализовать MCP requests через очередь
  - корректно завершать owned Excel instance и COM references
  - возвращать только сериализуемые Python values/DTO
  - обрабатывать COM busy/rejected calls без дублирования мутаций
extra_context: |
  Все COM proxies создаются, читаются и освобождаются в одном STA worker. MCP requests ставятся
  в очередь; параллельные команды Excel сериализуются. Инициализация — CoInitializeEx с
  COINIT_APARTMENTTHREADED; создание — DispatchEx("Excel.Application"); завершение — закрытие
  owned workbooks, Quit и освобождение ссылок в finally.
  ARCH-009 (binding): sync MCP tools выполняются SDK через anyio worker threads; поэтому api/* НЕ
  трогают COM, а только валидируют DTO и await ExcelComWorker.submit(). submit() должен быть
  awaitable с bounded timeout, без блокирующего join() в event loop; зависший worker → running/needs_recovery,
  а не stall всего сервера. ARCH-005: worker owned by MCPServer(lifespan=...) / main(); shutdown
  закрывает только owned workbooks → Quit → release → CoUninitialize в finally. ARCH-013: policy
  видимости — Visible=True для execute/compile/macro сессий, Visible=False только для inventory без UIA;
  DisplayAlerts менять только через context manager с восстановлением.
files:
  - path: src/excel_vba_mcp/adapters/excel/com_worker.py
    targets:
      - {type: class, name: ExcelComWorker}
      - {type: method, name: start}
      - {type: method, name: submit}
      - {type: method, name: shutdown}
  - path: src/excel_vba_mcp/adapters/excel/app_adapter.py
    targets: [{type: class, name: ExcelAppAdapter}]
changes:
  - action: implement_worker
    description: Implement the specified single-STA worker, queue, COM init/uninit and DispatchEx lifecycle.
  - action: add_error_mapping
    description: Map HRESULT and busy/rejected conditions to documented domain errors.
acceptance_criteria:
  - all Excel COM access executes on the worker thread
  - no COM proxy crosses into MCP/API layers
  - worker serializes requests and preserves operation order
  - shutdown closes only the Excel instance owned by the worker
  - COM busy/rejected operations are not blindly replayed
tests_to_run: ["unit tests for queue ordering and shutdown", "integration tests against isolated Excel instance"]
agents:
  - researcher: required_before_implementation; examine STA/COM/message-pump requirements
  - auditor: required; inspect worker lifecycle, ownership and deadlocks
  - planner: required; specify queue/shutdown/error matrix
  - validator: required_after_research_and_planner; review COM boundary and no-proxy-leak invariant
verification_mode: dedicated_gate
```

## task_011 — Verify COM worker lifecycle and isolation

```yaml
id: task_011
title: Verify STA worker serialization and Excel process lifecycle
priority: critical
depends_on: [task_010]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "5.3. COM и параллелизм; 11. Надёжность; 12. A02, A14"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A02 — Excel isolation", "A14 — Recovery"]
description: >
  Независимо проверить production worker на Excel: сериализацию, отсутствие перехвата
  пользовательской сессии, COM proxy lifetime и штатное завершение.
goals:
  - проверить несколько MCP requests при последовательном Excel execution
  - проверить закрытие owned workbooks и owned Excel instance
  - проверить busy/rejected path без повтора mutating call
  - проверить состояние worker после COM exception
extra_context: |
  При COM exception или modal window сервер не запускает параллельную Excel операцию и
  оставляет диагностический след.
files:
  - path: tests/integration_excel/test_com_worker.py
    targets:
      - {type: test_module, name: COM worker integration tests}
      - {type: test_case, name: test_requests_are_serialized}
      - {type: test_case, name: test_owned_excel_instance_shutdown}
      - {type: test_case, name: test_com_busy_does_not_replay_mutation}
changes:
  - action: add_integration_tests
    description: Проверить worker против live target Excel process и documented failure paths.
acceptance_criteria:
  - A02 and relevant A14 worker checks pass
  - existing user Excel session remains untouched
  - no concurrent COM call starts while worker handles a modal/busy state
tests_to_run: ["test_com_worker.py on target Windows/Excel build"]
agents:
  - auditor: required; review independent test coverage of COM ownership/recovery
  - validator: required; verify no hidden parallel COM path remains
verification_mode: dedicated_verification
```

## task_012 — Implement workbook session lifecycle and base tools

```yaml
id: task_012
title: Implement workbook session lifecycle and base MCP tools
priority: high
depends_on: [task_008, task_009, task_010, task_007]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "6.1. Базовые инструменты книги и сессии; 7.1. Workbook / Excel Session"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать open/close/save/status/list workbook operations через выделенную Excel session.
  Использовать allowlist, inspect/execute modes и явное управление application state.
goals:
  - реализовать excel_open, excel_close, excel_save, excel_get_status, excel_list_workbooks
  - поддержать inspect и execute с разными AutomationSecurity settings
  - явно задавать UpdateLinks, ReadOnly, Notify и AddToMru
  - работать с рабочими копиями в work/
  - восстанавливать временные Application settings в finally
  - поддержать lock по canonical path + workbook_id
extra_context: |
  По умолчанию работать с копией в work/; original path и actual working path должны быть видны.
  Excel открывается в отдельном DispatchEx instance. Inspection: AutomationSecurity=ForceDisable.
  Execution: пользователь/политика предварительно доверяет файлу; ByUI непосредственно перед
  Workbooks.Open. Close без явного save_changes не должен сохранять изменения.
files:
  - path: src/excel_vba_mcp/application/session_service.py
    targets:
      - {type: class, name: SessionService}
      - {type: method, name: open_workbook}
      - {type: method, name: close_workbook}
      - {type: method, name: save_workbook}
      - {type: method, name: get_status}
  - path: src/excel_vba_mcp/adapters/excel/workbook_adapter.py
    targets: [{type: class, name: WorkbookAdapter}]
  - path: src/excel_vba_mcp/api/workbook_tools.py
    targets:
      - {type: function, name: excel_open}
      - {type: function, name: excel_close}
      - {type: function, name: excel_save}
      - {type: function, name: excel_get_status}
      - {type: function, name: excel_list_workbooks}
  - path: src/excel_vba_mcp/server.py
    targets: [{type: function, name: workbook_tools registration}]
changes:
  - action: add_service_operations
    description: Реализовать workbook/session operations через SessionService и COM worker; не помещать COM logic в MCP handlers.
  - action: add_policy_checks
    description: Применять PathPolicy, MacroPolicy and ApprovalPolicy до open/save.
acceptance_criteria:
  - only allowlisted workbooks are opened
  - inspect mode does not enable workbook macros
  - execute mode requires trusted working copy and approved context
  - close without explicit save_changes never saves silently
  - status returns session/workbook identity, busy/read-only/dirty state and active operation
  - temporary Application settings are restored on success and exception
tests_to_run: ["unit tests for argument/state validation", "integration tests for inspect/execute open, close and explicit save"]
agents:
  - planner: recommended; review session state transitions and exception cleanup
  - validator: recommended; verify inspect/execute separation and close semantics
verification_mode: inline
```

## task_013 — Verify workbook session security and lifecycle

```yaml
id: task_013
title: Verify workbook session modes and safe close semantics
priority: high
depends_on: [task_012]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.1. Workbook / Excel Session; 10.1. Безопасный порядок открытия; 12. A01, A02, A15"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A01 — Запуск", "A02 — Excel isolation", "A15 — Security"]
description: >
  Отдельно проверить open/close flows с контролем пути, AutomationSecurity, исходной книги
  и явного сохранения.
goals:
  - проверить inspect/execute на fixtures
  - проверить отказ для запрещённого пути
  - проверить close с save_changes false/true
  - проверить возврат Application settings к исходному состоянию
extra_context: |
  Application settings восстанавливаются в finally. Trust Center, MacroSettings и Trusted
  Locations сервер автоматически не меняет.
files:
  - path: tests/integration_excel/test_workbook_session.py
    targets: [{type: test_module, name: workbook session integration tests}]
changes:
  - action: add_integration_tests
    description: Независимо проверить видимое open/close/save behavior и security invariants.
acceptance_criteria:
  - relevant A01, A02 and A15 checks pass
  - original input workbook remains unchanged in default flow
  - no workbook is saved without an explicit save instruction
tests_to_run: ["test_workbook_session.py on target Excel build"]
agents:
  - auditor: recommended; inspect security assertions
  - validator: recommended; independently verify session evidence
verification_mode: dedicated_verification
```

## task_014 — Implement immutable snapshots and restore

```yaml
id: task_014
title: Implement immutable workbook snapshots and restore
priority: critical
depends_on: [task_006, task_012, task_007, task_008]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "9.1. Snapshot model; 6.4. Testing & Snapshots"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать snapshot service поверх SaveCopyAs для открытой книги и file copy для закрытой.
  Сохранять неизменяемый snapshot, manifest, hash и metadata.
goals:
  - реализовать snapshot_create, snapshot_list, snapshot_restore, snapshot_diff
  - сохранять SHA-256, timestamp, source path, Excel version и model metadata
  - сохранять VBA component exports, если VBProject доступен
  - restore должен закрыть working copy без сохранения и открыть выбранный snapshot
  - SNAPSHOT_FAILED должен блокировать mutating operation
extra_context: |
  Snapshot содержит исходный файл, manifest.json, SHA-256, timestamp, исходный путь,
  версию Excel, workbook model и доступные VBA exports. Каждый test case запускается в новой
  disposable copy из baseline; предыдущий case не передаёт состояние следующему.
files:
  - path: src/excel_vba_mcp/application/snapshot_service.py
    targets:
      - {type: class, name: SnapshotService}
      - {type: method, name: create_snapshot}
      - {type: method, name: restore_snapshot}
      - {type: method, name: diff_snapshots}
  - path: src/excel_vba_mcp/storage/snapshot_store.py
    targets: [{type: class, name: SnapshotStore}]
  - path: src/excel_vba_mcp/api/test_tools.py
    targets:
      - {type: function, name: snapshot_create}
      - {type: function, name: snapshot_list}
      - {type: function, name: snapshot_restore}
      - {type: function, name: snapshot_diff}
changes:
  - action: add_snapshot_operations
    description: Реализовать open/closed workbook snapshot branches и manifest/hash contract.
  - action: enforce_pre_mutation_snapshot
    description: Требовать snapshot success до workbook/VBA mutation и test-case execution.
acceptance_criteria:
  - snapshot is immutable and contains manifest plus SHA-256
  - open workbook snapshot uses SaveCopyAs without replacing in-memory workbook
  - restore opens baseline copy without saving mutated copy
  - snapshot failure blocks relevant mutation
  - snapshot is not represented as rollback for external macro side effects
tests_to_run: ["unit tests for manifest/hash", "integration create/list/restore/diff on .xlsm fixture"]
agents:
  - researcher: required; verify SaveCopyAs and file-copy branch compatibility
  - auditor: required; inspect integrity and restore failures
  - planner: required; define snapshot invariants and pre-mutation ordering
  - validator: required_after_research_and_planner; review baseline identity and mutation gate
verification_mode: dedicated_gate
```

## task_015 — Verify snapshot integrity and isolation

```yaml
id: task_015
title: Verify snapshots, independent test copies and restore behavior
priority: critical
depends_on: [task_014]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "9.1. Snapshot model; 9.4. Restore, side effects and limits; 12. A09"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A09 — Snapshots"]
description: >
  Независимо проверить, что snapshot является воспроизводимым baseline и каждый test case
  получает независимую copy; проверить file integrity и restore scenarios.
goals:
  - проверить SHA-256 до и после test mutation
  - проверить отсутствие переноса изменений между двумя cases
  - проверить restore после обычного завершения и ошибки Excel
  - проверить, что внешние side effects не объявляются отменёнными
extra_context: |
  Snapshot не отменяет внешние эффекты макроса: БД, email, внешние файлы, сети, другие книги,
  печать и т.п. Snapshot не является sandbox.
files:
  - path: tests/integration_excel/test_snapshot_service.py
    targets:
      - {type: test_module, name: snapshot service integration tests}
      - {type: test_case, name: test_each_case_starts_from_same_baseline}
      - {type: test_case, name: test_restore_after_mutation}
changes:
  - action: add_independent_verification
    description: Проверить workbook-file identity и restore behavior разными test cases.
acceptance_criteria:
  - A09 passes
  - one test case cannot mutate baseline used by another case
  - failed restore returns structured error and keeps diagnostics
tests_to_run: ["test_snapshot_service.py on target Excel build"]
agents:
  - auditor: required; review isolation assumptions
  - validator: required; challenge baseline/hash assertions
verification_mode: dedicated_verification
```

## task_016 — Implement VBA inventory, procedure discovery and export

```yaml
id: task_016
title: Implement VBA Project inventory, procedure discovery and export
priority: critical
depends_on: [task_003, task_012, task_007]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "6.2. VBA Project; 7.2. VBA Project"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать read-only часть VBA Project Service: сведения о проекте, компонентах, процедурах,
  исходном коде, references и экспортируемых файлах. VBIDE — источник истины.
goals:
  - реализовать vba_project_info, vba_list_components, vba_list_procedures, vba_read_component
  - реализовать vba_export_project, vba_get_references, vba_search_code
  - различать Standard, Class, UserForm и Document Modules
  - извлекать процедуры через CodeModule API; regex fallback помечать heuristic
  - возвращать source_hash и revision metadata
extra_context: |
  Процедуры получать через VBIDE CodeModule API (ProcOfLine / ProcStartLine / ProcCountLines / Find),
  а не только regex. Standard/Class exports — .bas/.cls; UserForm — .frm + .frx. Document Modules
  не удалять/пересоздавать. Locked project → VBA_PROJECT_LOCKED; IsBroken references → missing_references.
files:
  - path: src/excel_vba_mcp/application/vba_project_service.py
    targets:
      - {type: class, name: VBAProjectService}
      - {type: method, name: get_project_info}
      - {type: method, name: list_components}
      - {type: method, name: list_procedures}
      - {type: method, name: read_component}
      - {type: method, name: export_project}
  - path: src/excel_vba_mcp/adapters/excel/vbide_adapter.py
    targets: [{type: class, name: VBIDEAdapter}]
  - path: src/excel_vba_mcp/api/vba_tools.py
    targets:
      - {type: function, name: vba_project_info}
      - {type: function, name: vba_list_components}
      - {type: function, name: vba_list_procedures}
      - {type: function, name: vba_read_component}
      - {type: function, name: vba_export_project}
      - {type: function, name: vba_get_references}
      - {type: function, name: vba_search_code}
changes:
  - action: add_read_only_vba_operations
    description: Использовать VBProject, VBComponents, CodeModule и References через Excel adapter.
  - action: add_component_manifest
    description: Сохранять component type, source hash, export status и UserForm resource metadata.
acceptance_criteria:
  - all component types are mapped correctly
  - UserForms retain .frm and .frx
  - CodeModule API is used for procedure discovery where available
  - fallback parsing is marked heuristic
  - missing references and locked/unavailable projects are structured
tests_to_run: ["unit tests for component/procedure DTOs", "integration tests with CleanProject and LockedProject fixtures"]
agents:
  - researcher: required; check VBIDE methods and component export behavior
  - auditor: required; inspect Document Module and UserForm resource handling
  - planner: required; prepare fixture matrix
  - validator: required_after_research_and_planner; review export fidelity and read-only guarantee
verification_mode: dedicated_gate
```

## task_017 — Verify VBA inventory and export fidelity

```yaml
id: task_017
title: Verify VBA component inventory and export fidelity
priority: high
depends_on: [task_016]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.2. VBA Project; 12.1. Test fixtures; 12. A03, A04"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A03 — Workbook model", "A04 — VBA read/export"]
description: >
  Независимо проверить полноту VBA inventory и round-trip exports на fixture set.
goals:
  - проверить Standard, Class, Document Module и UserForm
  - сравнить source hash с exported source
  - проверить references, IsBroken и locked project diagnostics
extra_context: |
  Не полагаться только на распаковку OOXML vbaProject.bin как на единственный источник истины.
  Не обходить пароль/защиту VBA проекта.
files:
  - path: tests/integration_excel/test_vba_project_read.py
    targets: [{type: test_module, name: VBA project read/export integration tests}]
changes:
  - action: add_independent_verification
    description: Проверить наблюдаемые component inventory, hashes и exported artifacts.
acceptance_criteria:
  - A03 and A04 pass for supported fixture components
  - source hash matches exported source representation
  - UserForm designer resource is retained
  - locked project is reported without bypass attempt
tests_to_run: ["test_vba_project_read.py against VBA fixtures"]
agents:
  - auditor: required; review coverage of all VBComponent types
  - validator: required; independently verify fixture assertions
verification_mode: dedicated_verification
```

## task_018 — Implement guarded VBA source mutations

```yaml
id: task_018
title: Implement snapshot-backed VBA source replacement and insertion
priority: critical
depends_on: [task_014, task_016, task_009]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "3.2. Основные потоки данных; 7.2. VBA Project"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать VBA source mutations только в рабочей copy, со snapshot до изменений и source revision check.
goals:
  - реализовать vba_replace_component_source для Standard/Class modules
  - реализовать vba_insert_code с семантической позицией, например декларационная секция или процедура
  - проверять expected source revision/hash до записи
  - экспортировать before/after source и hashes
  - не удалять и не пересоздавать Document Modules
extra_context: |
  Изменение кода — только в рабочей копии, после snapshot и проверки source revision, с журналом
  before/after hash. Вставку не выполнять по неоднозначной текстовой эвристике.
files:
  - path: src/excel_vba_mcp/application/vba_project_service.py
    targets:
      - {type: method, name: replace_component_source}
      - {type: method, name: insert_code}
  - path: src/excel_vba_mcp/api/vba_tools.py
    targets:
      - {type: function, name: vba_replace_component_source}
      - {type: function, name: vba_insert_code}
  - path: src/excel_vba_mcp/security/approval.py
    targets: [{type: method, name: authorize_vba_write}]
changes:
  - action: add_guarded_mutations
    description: Добавить source replace/insert через существующий VBAProjectService, SnapshotService, revision check и approval.
  - action: add_change_manifest
    description: Фиксировать before/after hash и exported source artifacts.
acceptance_criteria:
  - mutation is rejected for original/read-only workbooks
  - snapshot exists before mutation
  - stale revision is rejected without overwriting current source
  - before/after hash and audit event are returned
  - Document Modules are not removed/recreated
tests_to_run: ["unit tests for revision mismatch and semantic insertion", "integration tests on disposable VBA fixture copies"]
agents:
  - researcher: required; compare safe CodeModule editing operations within VBIDE constraints
  - auditor: required; review overwrite and Document Module risks
  - planner: required; sequence snapshot, approval, edit and export checks
  - validator: required_after_research_and_planner; review stale-write and rollback tests
verification_mode: dedicated_gate
```

## task_019 — Verify VBA source mutation safety

```yaml
id: task_019
title: Verify VBA source edits, revision conflicts and snapshots
priority: critical
depends_on: [task_018]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.2. VBA Project; 12. A05"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A05 — VBA edit"]
description: >
  Отдельно проверить, что VBA changes связаны со snapshot, approval и source revision.
goals:
  - проверить replace/insert на disposable module
  - проверить stale source hash conflict
  - проверить блокировку mutation при отсутствии snapshot/approval
  - проверить before/after exports и hashes
extra_context: |
  Операции изменения кода разрешены только на рабочей copy после snapshot, source revision
  verification и с журналом before/after hash.
files:
  - path: tests/integration_excel/test_vba_project_write.py
    targets: [{type: test_module, name: VBA project mutation tests}]
changes:
  - action: add_independent_verification
    description: Проверить mutation safety and provenance на disposable fixtures.
acceptance_criteria:
  - A05 passes
  - stale revision never overwrites current code
  - failed mutation leaves snapshot and diagnostics available
tests_to_run: ["test_vba_project_write.py"]
agents:
  - auditor: required; review failure/recovery scenarios
  - validator: required; independently inspect source revision assertions
verification_mode: dedicated_verification
```

## task_020 — Implement VBE compile service and UIA diagnostics

```yaml
id: task_020
title: Implement the VBE compile service and diagnostic capture
priority: critical
depends_on: [task_004, task_010, task_016, task_007]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "8. Компиляция VBA; 6.3. Компиляция и запуск"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать production Compile Service на базе PoC. Вызывать VBE Debug → Compile,
  координировать COM и UIA и возвращать только доказуемый status.
goals:
  - реализовать vba_compile и vba_get_compile_diagnostics
  - координировать VBE command invocation и UIA watcher
  - поддержать compiled, compile_error, already_compiled и indeterminate
  - сохранять dialog text, module/context, selected token/context, screenshot и fallback notes
  - собирать missing references до/после compile
  - блокировать последующие COM actions до разрешения compile modal
extra_context: |
  Сначала искать Compile control по доступной локализованной подписи/структуре меню; control ID
  использовать как fallback только после проверки наличия и ожидаемой подписи. UIA watcher
  запускается до compile. Нельзя сообщать compiled только из-за отсутствия COM exception.
files:
  - path: src/excel_vba_mcp/application/compile_service.py
    targets:
      - {type: class, name: CompileService}
      - {type: method, name: compile_project}
      - {type: method, name: get_last_diagnostics}
  - path: src/excel_vba_mcp/adapters/windows_ui/vbe_compile_ui.py
    targets: [{type: class, name: VBECompileUI}]
  - path: src/excel_vba_mcp/adapters/windows_ui/dialog_reader.py
    targets: [{type: class, name: DialogReader}]
  - path: src/excel_vba_mcp/api/compile_tools.py
    targets:
      - {type: function, name: vba_compile}
      - {type: function, name: vba_get_compile_diagnostics}
changes:
  - action: add_compile_service
    description: Реализовать подтверждённый PoC путь вызова Compile; UI selectors остаются в windows_ui adapter.
  - action: add_compile_diagnostics
    description: Возвращать structured statuses и evidence; при неуверенности сохранять indeterminate.
acceptance_criteria:
  - Compile invokes the VBE project compile command, not a macro
  - compile-error fixture produces COMPILE_ERROR with evidence
  - unverified command/UIA state produces COMPILE_INDETERMINATE
  - already_compiled is not described as a new compile run
  - next COM action is blocked while compile dialog remains unresolved
tests_to_run: ["unit tests for status/diagnostic DTOs", "integration success/error fixtures on target Office build"]
agents:
  - researcher: required; compare command/UIA routes validated by PoC
  - auditor: required; inspect false-success and modal-dialog handling
  - planner: required; define command lookup, watcher and recovery sequence
  - validator: required_after_research_and_planner; independently review compile status claims
verification_mode: dedicated_gate
```

## task_021 — Verify compile outcomes and diagnostics

```yaml
id: task_021
title: Verify VBA compile outcome classification and diagnostics
priority: critical
depends_on: [task_020]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "8.1–8.3. Compile mechanism and critical nuances; 12. A06, A07"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A06 — Compile success", "A07 — Compile error"]
description: >
  Независимо проверить compile success, compile error и неопределённый UI state.
  Главная цель — исключить ложный compiled.
goals:
  - проверить compile status для чистого проекта
  - проверить error message и выделенный контекст
  - проверить indeterminate при недоступном VBE/UIA
  - проверить отсутствие зависшего modal dialog перед следующей COM operation
extra_context: |
  Компиляция зависит от VBE UI, Office version/locale, dialogs и UI Automation. Универсальный
  headless compile не обещается.
files:
  - path: tests/integration_excel/test_vba_compile.py
    targets: [{type: test_module, name: VBE compile verification tests}]
changes:
  - action: add_independent_verification
    description: Проверить все compile outcomes на positive и negative fixtures.
acceptance_criteria:
  - A06 and A07 pass without false-positive compiled statuses
  - COMPILE_INDETERMINATE is returned for unconfirmed outcomes
  - VBE dialog is handled before Excel COM is used again
tests_to_run: ["test_vba_compile.py on each supported Office build"]
agents:
  - auditor: required; inspect negative/false-positive cases
  - validator: required; review diagnostic evidence against acceptance claims
verification_mode: dedicated_verification
```

## task_022 — Implement policy-controlled macro execution

```yaml
id: task_022
title: Implement policy-controlled macro execution service
priority: critical
depends_on: [task_005, task_009, task_016, task_020]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "6.3. Компиляция и запуск; 7.3. Macro Execution"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать macro_list, macro_run и macro_get_last_result. Запускать только найденную/разрешённую
  процедуру по qualified workbook/procedure name с policy approval.
goals:
  - перечислять кандидатов на запуск из VBA project inventory
  - проверять имя, доступность процедуры и policy
  - вызывать Application.Run с позиционными сериализуемыми аргументами
  - возвращать run_id, timestamps, duration, return_value, status и structured error
  - не повторять запуск после timeout/COM busy без доказательства, что макрос не стартовал
  - в test-run flow требовать compile success/already_compiled до macro execution
extra_context: |
  Запускать только явно указанную процедуру, обнаруженную в текущем VBA project либо вручную
  разрешённую policy. Не принимать имя макроса как свободный невалидированный текст.
  Application.Run использует полное квалифицированное имя книги и процедуры и позиционные args.
  Execution требует approval/allowlist по workbook path/hash + macro procedure; запуск логируется.
files:
  - path: src/excel_vba_mcp/application/macro_service.py
    targets:
      - {type: class, name: MacroService}
      - {type: method, name: list_macros}
      - {type: method, name: run_macro}
      - {type: method, name: get_last_result}
  - path: src/excel_vba_mcp/api/macro_tools.py
    targets:
      - {type: function, name: macro_list}
      - {type: function, name: macro_run}
      - {type: function, name: macro_get_last_result}
changes:
  - action: add_macro_operations
    description: Реализовать macro discovery, policy gate, Application.Run и result normalization через COM worker.
  - action: add_run_audit
    description: Логировать macro identity, workbook hash, approval, duration, status и HRESULT/diagnostics.
acceptance_criteria:
  - unapproved or unknown macros are blocked
  - macro uses qualified workbook/procedure name
  - supported args are validated and positional
  - return value is serialized when possible
  - timeout/runtime errors do not cause blind retries
  - test-run execution checks compile result first
tests_to_run: ["unit tests for name/argument/policy", "integration tests for successful/missing/runtime-error macros"]
agents:
  - researcher: required; inspect Application.Run behavior and runtime failure observability
  - auditor: required; review approval, timeout and modal-error handling
  - planner: required; prepare execution/result contract and failure matrix
  - validator: required_after_research_and_planner; review no-retry and approval guarantees
verification_mode: dedicated_gate
```

## task_023 — Verify macro execution and safety

```yaml
id: task_023
title: Verify macro execution results, errors and approval policy
priority: critical
depends_on: [task_022]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.3. Macro Execution; 11.1. Ограничения; 12. A08"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A08 — Macro run"]
description: >
  Независимо проверить macro list/run/result и поведение разрешённых/запрещённых процедур.
goals:
  - проверить успешный public fixture macro
  - проверить blocked и not-found procedures
  - проверить runtime error и modal detection
  - проверить, что timeout не вызывает автоматический повтор
extra_context: |
  Application.Run может показать runtime error dialog; при automation caller может получить
  общий error вместо исходного VBA error number/description. Нельзя обещать точную диагностику,
  если Excel её не передаёт.
files:
  - path: tests/integration_excel/test_macro_execution.py
    targets: [{type: test_module, name: macro execution integration tests}]
changes:
  - action: add_independent_verification
    description: Проверить наблюдаемые результаты macro execution и policy decisions.
acceptance_criteria:
  - A08 passes
  - blocked macro never runs
  - successful return value and duration are captured
  - runtime error is classified using available evidence without false precision
tests_to_run: ["test_macro_execution.py on target Excel build"]
agents:
  - auditor: required; verify error classification and policy coverage
  - validator: required; review runtime outcome evidence
verification_mode: dedicated_verification
```

## task_024 — Implement Sheets inventory and operations

```yaml
id: task_024
title: Implement Sheets inventory and worksheet operations
priority: high
depends_on: [task_012, task_007, task_008, task_009, task_014]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.5. Sheets; 6.5. Shapes, Controls, Sheets, Tables, Names"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать Phase 1 Sheets inventory и операции над Worksheets, различая коллекции Sheets
  и Worksheets. CodeName возвращать только как read-only identifier.
goals:
  - реализовать sheets_list и sheet_get
  - реализовать sheet_add, sheet_rename, sheet_copy, sheet_delete, sheet_set_visibility
  - возвращать name, index, object_type, visibility и для Worksheet CodeName/protected/table count/UsedRange metadata
  - требовать snapshot и approval для delete
  - не удалять последний видимый лист
extra_context: |
  Инвентаризация через Workbook.Sheets (все типы) и Workbook.Worksheets (только Worksheet).
  Создание в MVP — Worksheet по умолчанию. Поддержать rename, copy, delete и Visible.
  Удаление — snapshot + явное подтверждение + проверка, что остаётся хотя бы один видимый лист.
  CodeName не менять.
files:
  - path: src/excel_vba_mcp/application/sheets_service.py
    targets:
      - {type: class, name: SheetsService}
      - {type: method, name: list_sheets}
      - {type: method, name: add_sheet}
      - {type: method, name: delete_sheet}
  - path: src/excel_vba_mcp/adapters/excel/sheets_adapter.py
    targets: [{type: class, name: SheetsAdapter}]
  - path: src/excel_vba_mcp/api/sheet_tools.py
    targets:
      - {type: function, name: sheets_list}
      - {type: function, name: sheet_get}
      - {type: function, name: sheet_add}
      - {type: function, name: sheet_rename}
      - {type: function, name: sheet_copy}
      - {type: function, name: sheet_delete}
      - {type: function, name: sheet_set_visibility}
changes:
  - action: add_sheets_operations
    description: Реализовать metadata and worksheet operations через Excel Object Model.
acceptance_criteria:
  - Sheets and Worksheets are distinct in returned model
  - CodeName is returned but never changed
  - delete requires snapshot and approval
  - delete is rejected if no visible worksheet would remain
  - postconditions reflect actual collection state
tests_to_run: ["unit tests for metadata/preconditions", "integration tests for worksheet/non-worksheet cases"]
agents:
  - auditor: recommended; inspect delete/visibility safeguards
  - validator: recommended; review destructive-operation tests
verification_mode: inline
```

## task_025 — Verify Sheets operations

```yaml
id: task_025
title: Verify Sheets inventory and destructive-operation safeguards
priority: high
depends_on: [task_024]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.5. Sheets; 12. A03, A11"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A11 — Sheets"]
description: >
  Отдельно проверить Sheets/Worksheets inventory и фактическое изменение состава листов.
goals:
  - проверить worksheet metadata и CodeName
  - проверить add/rename/copy/visibility
  - проверить delete approval, snapshot и visible-sheet guard
extra_context: |
  Не менять CodeName: это свойство VBIDE/document module с отдельными последствиями.
files:
  - path: tests/integration_excel/test_sheets_service.py
    targets: [{type: test_module, name: Sheets service integration tests}]
changes:
  - action: add_independent_verification
    description: Проверить user-visible sheet behavior и destructive-operation controls.
acceptance_criteria:
  - A11 passes
  - supported non-worksheet sheet types are represented truthfully
  - deletion without snapshot/approval or leaving no visible sheet is blocked
tests_to_run: ["test_sheets_service.py"]
agents:
  - auditor: recommended; review sheet collection coverage
  - validator: recommended; verify acceptance evidence
verification_mode: dedicated_verification
```

## task_026 — Implement ListObject inventory and row/column operations

```yaml
id: task_026
title: Implement Excel Table inventory and ListObject mutations
priority: critical
depends_on: [task_012, task_007, task_008, task_009, task_014]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.6. Tables (ListObjects); 6.5. Shapes, Controls, Sheets, Tables, Names"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать Excel Tables через ListObject, ListRows и ListColumns. Мутации должны адресовать
  однозначные строки, выполняться после snapshot и approval и проверяться postconditions.
goals:
  - реализовать tables_list, table_get, table_read_rows
  - реализовать table_append_rows, table_insert_row, table_update_rows, table_delete_rows
  - реализовать table_add_column, table_delete_column, table_create, table_resize
  - возвращать Range/HeaderRowRange/DataBodyRange/TotalRowRange/ShowTotals/TableStyle metadata
  - поддержать row offset/limit и response size limit
  - сохранять типы дат/чисел/boolean и типизировать Excel cell errors
  - проверять counts, range boundaries, inserted values и уникальность заголовков
  - возвращать preview/error при неоднозначном update/delete
extra_context: |
  Таблица Excel — ListObject, не просто range. Добавление/удаление строк/столбцов выполняется
  ListRows/ListColumns, чтобы сохранить table structure, name, filters и связанные свойства.
  Строки добавляются через ListRows.Add; столбцы — через ListColumns.Add(Position) с уникальным
  заголовком. Ambiguous updates/deletes не выполняются автоматически.
files:
  - path: src/excel_vba_mcp/application/tables_service.py
    targets:
      - {type: class, name: TablesService}
      - {type: method, name: list_tables}
      - {type: method, name: read_table_rows}
      - {type: method, name: append_table_rows}
      - {type: method, name: update_table_rows}
      - {type: method, name: add_table_column}
  - path: src/excel_vba_mcp/adapters/excel/tables_adapter.py
    targets: [{type: class, name: TablesAdapter}]
  - path: src/excel_vba_mcp/api/table_tools.py
    targets:
      - {type: function, name: tables_list}
      - {type: function, name: table_get}
      - {type: function, name: table_read_rows}
      - {type: function, name: table_append_rows}
      - {type: function, name: table_insert_row}
      - {type: function, name: table_update_rows}
      - {type: function, name: table_delete_rows}
      - {type: function, name: table_add_column}
      - {type: function, name: table_delete_column}
      - {type: function, name: table_create}
      - {type: function, name: table_resize}
changes:
  - action: add_listobject_operations
    description: Реализовать Stage 1 ListObject tools с snapshot-before-mutation, deterministic row addressing и postcondition checks.
acceptance_criteria:
  - rows/columns are changed through ListRows/ListColumns, not adjacent cell writes
  - metadata includes headers, body, totals and range information
  - supported Excel types are preserved; cell errors are typed
  - ambiguous updates/deletes do not mutate data
  - postconditions verify counts, boundaries and inserted values
tests_to_run: ["unit tests for key resolution/conversion/ambiguity", "integration tests with TableNamesProject.xlsm"]
agents:
  - researcher: required; check ListObject edge cases including totals/formulas/insertion positions
  - auditor: required; inspect mutation and ambiguity safeguards
  - planner: required; define fixtures for empty tables, totals and formula columns
  - validator: required_after_research_and_planner; review table-preservation checks
verification_mode: dedicated_gate
```

## task_027 — Verify ListObject mutation semantics

```yaml
id: task_027
title: Verify table row and column operations preserve ListObject structure
priority: critical
depends_on: [task_026]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.6. Tables (ListObjects); 12. A12"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A12 — Tables"]
description: >
  Независимо проверить row/column mutations на таблицах с обычным и пустым телом, totals row,
  formula columns и неоднозначными ключами.
goals:
  - проверить ListObject identity/name и расширение Range
  - проверить types, totals и formula behavior
  - проверить ambiguous update/delete preview without mutation
  - проверить snapshot/approval gates
extra_context: |
  После изменений проверять число строк/столбцов, границы Range, значения добавленной строки
  и отсутствие конфликтующих заголовков.
files:
  - path: tests/integration_excel/test_tables_service.py
    targets: [{type: test_module, name: ListObject integration tests}]
changes:
  - action: add_independent_verification
    description: Проверить реальную структуру таблицы и данные после каждого сценария.
acceptance_criteria:
  - A12 passes
  - table identity and structure remain correct after supported mutations
  - ambiguous row is never modified automatically
tests_to_run: ["test_tables_service.py on TableNamesProject.xlsm"]
agents:
  - auditor: required; inspect table edge-case assertions
  - validator: required; review no-adjacent-range-write guarantee
verification_mode: dedicated_verification
```

## task_028 — Implement workbook- and worksheet-scope Names

```yaml
id: task_028
title: Implement defined Names inventory and scoped CRUD
priority: high
depends_on: [task_012, task_007, task_008, task_009, task_014]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.7. Names; 6.5. Shapes, Controls, Sheets, Tables, Names"
source_file: excel_vba_mcp_technical_spec.docx
description: >
  Реализовать read/create/update/delete Defined Names с явным scope и корректной обработкой
  имён, которые не являются диапазонами.
goals:
  - реализовать names_list, name_get, name_create, name_update, name_delete
  - возвращать workbook/worksheet scope, NameLocal, Visible, RefersTo и RefersToLocal
  - вызывать RefersToRange в отдельной error-handling ветке
  - представлять range target и non-range target_kind без eval
  - обнаруживать scope/name conflicts
  - выполнять mutating operations после snapshot и approval
extra_context: |
  Names могут ссылаться на формулу, константу, динамическое выражение, внешнюю книгу или иметь
  ошибочную ссылку. Для range name возвращать sheet + address; для non-range — target_kind и
  RefersTo. Не вычислять формулу через eval; создавать/обновлять/удалять с явным scope.
files:
  - path: src/excel_vba_mcp/application/names_service.py
    targets:
      - {type: class, name: NamesService}
      - {type: method, name: list_names}
      - {type: method, name: get_name}
      - {type: method, name: create_name}
      - {type: method, name: update_name}
      - {type: method, name: delete_name}
  - path: src/excel_vba_mcp/adapters/excel/names_adapter.py
    targets: [{type: class, name: NamesAdapter}]
  - path: src/excel_vba_mcp/api/name_tools.py
    targets:
      - {type: function, name: names_list}
      - {type: function, name: name_get}
      - {type: function, name: name_create}
      - {type: function, name: name_update}
      - {type: function, name: name_delete}
changes:
  - action: add_defined_name_operations
    description: Реализовать scope-aware names operations и безопасную обработку range/non-range targets.
acceptance_criteria:
  - workbook and worksheet scopes are distinguished
  - non-range names are returned without forced RefersToRange conversion
  - formulas are never evaluated by Python
  - create/update/delete require explicit scope and conflict checks
  - mutations require snapshot and approval
tests_to_run: ["unit tests for scope conflict/non-range DTO", "integration tests with TableNamesProject.xlsm"]
agents:
  - researcher: required; review defined-name scope and non-range behavior
  - auditor: required; inspect external/invalid references and mutation cases
  - planner: required; define workbook/sheet scope fixtures
  - validator: required_after_research_and_planner; review non-range safety and conflicts
verification_mode: dedicated_gate
```

## task_029 — Verify Defined Names behavior

```yaml
id: task_029
title: Verify Defined Names scopes and non-range handling
priority: high
depends_on: [task_028]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.7. Names; 12. A13"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A13 — Names"]
description: >
  Отдельно проверить workbook/worksheet scope, range/non-range names и операции изменения/удаления.
goals:
  - проверить RefersToRange success и error branches
  - проверить RefersTo для формульных/константных/ошибочных имён
  - проверить scope conflict и сохранение области действия при update
extra_context: |
  Сохранять scope, имя, NameLocal (если доступно), Visible, RefersTo и RefersToLocal.
files:
  - path: tests/integration_excel/test_names_service.py
    targets: [{type: test_module, name: Defined Names integration tests}]
changes:
  - action: add_independent_verification
    description: Проверить поведение Excel Names collection и non-range result handling.
acceptance_criteria:
  - A13 passes
  - no non-range name causes unhandled error or forced range cast
  - mutations preserve requested scope
tests_to_run: ["test_names_service.py on TableNamesProject.xlsm"]
agents:
  - auditor: recommended; review edge-case fixture coverage
  - validator: recommended; independently review scope assertions
verification_mode: dedicated_verification
```

## task_030 — Implement Shapes and Controls inventory

```yaml
id: task_030
title: Implement Shapes and Controls inventory and macro associations
priority: high
depends_on: [task_009, task_012, task_014, task_016]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.6. Shapes & Controls; 8.4. ControlsService; 12.1. Test fixtures"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A10 — Shapes and Controls", "ControlsProject.xlsm fixture requirements"]
description: >
  Реализовать получение сведений о Shapes, Form Controls и ActiveX OLEObjects, а также
  установление документированных или вероятных связей с VBA-процедурами. Не реализовывать
  автоматическое нажатие элементов управления или UI-click тестирование.
goals:
  - перечислять фигуры и элементы управления выбранного листа/книги
  - различать обычные Shapes, Form Controls и ActiveX OLEObjects
  - возвращать имя, тип, подпись, лист, положение, видимость и доступные свойства
  - извлекать OnAction для поддерживающих его объектов
  - для ActiveX возвращать доступные метаданные и probable event procedure mapping
  - безопасно изменять OnAction только в пределах описанных операций и после snapshot/approval
extra_context: |
  Form Controls и фигуры могут связываться с макросами через OnAction. ActiveX находится
  в коллекции OLEObjects и использует обработчики событий в Document Module. Для ActiveX
  сопоставление имени контрола с процедурой события является вероятным, а не гарантированным;
  результат обязан маркироваться как inferred/probable. Этап 1 не включает имитацию кликов,
  полную автоматизацию пользовательского интерфейса ActiveX или изменение исходников событий.
files:
  - path: src/excel_vba_mcp/application/controls_service.py
    targets:
      - {type: class, name: ControlsService}
      - {type: method, name: list_controls}
      - {type: method, name: get_control}
      - {type: method, name: get_macro_association}
      - {type: method, name: set_on_action}
  - path: src/excel_vba_mcp/adapters/excel/controls_adapter.py
    targets: [{type: class, name: ControlsAdapter}]
  - path: src/excel_vba_mcp/api/shape_tools.py
    targets:
      - {type: function, name: shapes_list}
      - {type: function, name: shape_get}
      - {type: function, name: shape_set_on_action}
  - path: src/excel_vba_mcp/api/control_tools.py
    targets:
      - {type: function, name: controls_list}
      - {type: function, name: control_get}
      - {type: function, name: control_set_on_action}
changes:
  - action: add_controls_inventory
    description: Добавить read-only перечисление Shapes, Form Controls и ActiveX OLEObjects.
  - action: add_macro_association_resolution
    description: Возвращать OnAction и отдельно маркировать эвристическое сопоставление ActiveX событий.
  - action: add_guarded_on_action_update
    description: Изменять OnAction только после snapshot, policy check и approval.
acceptance_criteria:
  - Form Controls, Shapes и ActiveX различаются в DTO
  - OnAction возвращается без запуска назначенного макроса
  - ActiveX event mapping обозначается как probable/inferred
  - отсутствующие или недоступные свойства возвращаются как диагностируемые значения, а не unhandled COM errors
  - изменение OnAction требует разрешённого рабочего файла, snapshot и approval
  - A10 inventory and association requirements are satisfied
tests_to_run: ["unit tests for control type mapping and association confidence", "integration tests with ControlsProject.xlsm"]
agents:
  - researcher: required; compare supported Shapes/Form Controls/ActiveX enumeration and event-mapping approaches
  - auditor: required; inspect ActiveX and document-module boundaries
  - planner: required; define fixture cases and result confidence model
  - validator: required_after_research_and_planner; verify no unsupported click-testing promise
verification_mode: dedicated_gate
```

## task_031 — Verify Shapes and Controls behavior

```yaml
id: task_031
title: Verify Shapes and Controls inventory and macro associations
priority: high
depends_on: [task_030]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.6. Shapes & Controls; 12.1. Test fixtures; 12.2. Acceptance criteria A10"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A10 — Shapes and Controls"]
description: >
  Независимо проверить типизацию элементов управления, чтение назначений макросов и
  маркировку эвристических ActiveX event mappings на реальной книге Excel.
goals:
  - покрыть Shape/Form Control/ActiveX объекты на fixture
  - проверить, что инвентаризация не запускает обработчики
  - проверить назначение/снятие OnAction в тестовой копии
  - проверить недоступные и удалённые целевые процедуры
extra_context: |
  Проверка проводится на disposable workbook copy. Не нажимать ActiveX/Form Controls
  программно; фактическое выполнение макроса проверяется отдельным Macro Execution сервисом.
files:
  - path: tests/integration_excel/test_controls_service.py
    targets: [{type: test_module, name: Shapes and Controls integration tests}]
changes:
  - action: add_independent_verification
    description: Проверить A10 и регрессионные случаи перечисления/связи Controls.
acceptance_criteria:
  - A10 passes on ControlsProject.xlsm
  - inventory is read-only and does not invoke control actions
  - inferred ActiveX mappings are never reported as confirmed
  - OnAction update is isolated to the disposable copy
  - failures are represented by structured diagnostics
tests_to_run: ["test_controls_service.py on ControlsProject.xlsm"]
agents:
  - auditor: recommended; review ActiveX fixture edge cases
  - validator: recommended; independently review A10 evidence
verification_mode: dedicated_verification
```

## task_032 — Aggregate the Stage 1 workbook model and read-only resources

```yaml
id: task_032
title: Aggregate Stage 1 workbook model and expose read-only MCP resources
priority: high
depends_on: [task_017, task_025, task_027, task_029, task_031]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.1. Workbook model; 6. MCP resources; 12.2. Acceptance criteria A03"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A03 — Workbook model", "MCP resources and discovery"]
description: >
  Собрать результаты уже реализованных inventory services в единый Stage 1 workbook model.
  Модель должна быть агрегирующим представлением, а не вторым источником объектных данных.
goals:
  - получать идентичность книги и состав доступных объектов этапа 1
  - объединять sheets, tables, names, VBA components/procedures, Shapes и Controls
  - включать диагностические статусы недоступных или защищённых компонентов
  - предоставлять read-only MCP resource для полной/выборочной модели
  - предоставлять read-only ресурсы для списка VBA components и исходного кода выбранного компонента
  - избегать загрузки всех ячеек/диапазонов в модель
extra_context: |
  Model строится через WorkbookModelService и использует DTO доменного слоя. На этом этапе
  регистрируются ресурсы workbook model и VBA component/code. Ресурс результатов тестовых
  запусков добавляется вместе с RunRepository в задаче task_035. При недоступном VBProject
  модель должна сохранять информацию о книге и возвращать отдельный VBA access status.
files:
  - path: src/excel_vba_mcp/application/workbook_model_service.py
    targets:
      - {type: class, name: WorkbookModelService}
      - {type: method, name: get_workbook_model}
      - {type: method, name: get_sheet_model}
  - path: src/excel_vba_mcp/domain/models.py
    targets: [{type: model_group, name: workbook/sheet/table/name/VBA/control model DTOs}]
  - path: src/excel_vba_mcp/api/workbook_tools.py
    targets: [{type: function, name: workbook_get_model}]
  - path: src/excel_vba_mcp/server.py
    targets: [{type: resource_registration, name: workbook model and VBA source resources}]
changes:
  - action: aggregate_existing_services
    description: Собрать модель из уже существующих services, не дублируя COM traversal в API handlers.
  - action: register_read_only_resources
    description: Зарегистрировать workbook model, VBA component inventory и component source resources.
acceptance_criteria:
  - модель содержит сведения о книгах, листах, таблицах, именах, VBA и элементах управления этапа 1
  - недоступность VBProject не прерывает получение остальной модели
  - resource чтения кода не изменяет книгу и не запускает VBA
  - в модель не включаются все ячейки и не добавляется Range/Cells API этапа 2
  - A03 Stage 1 model requirements are satisfied
tests_to_run: ["unit tests for model aggregation and unavailable component states", "integration test for full workbook model fixture"]
agents:
  - researcher: recommended; review MCP resource patterns against the selected SDK version
  - auditor: required; inspect consistency and duplicate COM traversal across services
  - planner: required; define aggregation order and partial-model behavior
  - validator: required_after_research_and_planner; review Stage 1 boundary and A03 coverage
verification_mode: dedicated_gate
```

## task_033 — Verify workbook model aggregation and resources

```yaml
id: task_033
title: Verify Stage 1 workbook model and read-only resources
priority: high
depends_on: [task_032]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.1. Workbook model; 6. MCP resources; 12.2. Acceptance criteria A03"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A03 — Workbook model"]
description: >
  Проверить полноту и согласованность workbook model на fixture книгах, включая частичную
  доступность при отключённом AccessVBOM и read-only поведение MCP resources.
goals:
  - сопоставить model inventory с непосредственными результатами соответствующих services
  - проверить идентификаторы и имена объектов
  - проверить partial model при недоступном/защищённом VBA project
  - проверить, что resource чтения кода соответствует VBIDE source
extra_context: |
  Использовать CleanProject.xlsm, ControlsProject.xlsm и fixture с недоступным VBA access,
  если окружение позволяет подготовить его без изменения глобальных Trust Center настроек.
files:
  - path: tests/integration_excel/test_workbook_model_service.py
    targets: [{type: test_module, name: Stage 1 workbook model integration tests}]
changes:
  - action: add_independent_verification
    description: Добавить тесты A03 и согласованности между агрегатором и сервисами объектов.
acceptance_criteria:
  - A03 passes for full-access fixture
  - partial model is returned when VBA access is unavailable
  - resource reads do not change workbook state
  - aggregated counts/names agree with their source services
tests_to_run: ["test_workbook_model_service.py on Stage 1 fixtures"]
agents:
  - auditor: recommended; review consistency and partial model cases
  - validator: recommended; independently verify A03 coverage
verification_mode: dedicated_verification
```

## task_034 — Harden operation recovery and artifact storage

```yaml
id: task_034
title: Harden Excel operation recovery and artifact storage
priority: critical
depends_on: [task_013, task_015, task_019, task_021, task_023, task_025, task_027, task_029, task_031, task_033]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "5.5. Error handling; 5.6. Recovery; 9. Security; 10. Logging; 11. Artifacts"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["Error handling and recovery requirements", "Security requirements", "Artifact storage requirements"]
description: >
  Завершить общие механизмы структурированных ошибок, восстановления состояния Excel,
  безопасного завершения сессии и сохранения test/compile/macro evidence. Не реализовывать
  test runner в этой задаче.
goals:
  - унифицировать диагностические статусы и отображение COM/UIA exceptions
  - обеспечить восстановление/закрытие owned Excel instance после ошибок
  - сохранять скриншоты, логи, экспорт кода и другие артефакты в разрешённой папке
  - применять path policy к путям артефактов
  - исключить секреты и чувствительные данные из логов
  - проверять восстановление worker/session после прерывания операции
extra_context: |
  Snapshot не является sandbox и не откатывает внешние side effects макросов. Не добавлять
  Windows Service, удалённый транспорт, произвольный shell/COM endpoint или механизм обхода
  макро-защиты. ArtifactStore хранит evidence, а не выполняет логику тестов. Восстановление
  ограничено тем, что спецификация обещает для Excel workbook/session state.
files:
  - path: src/excel_vba_mcp/domain/errors.py
    targets: [{type: module, name: structured error taxonomy}]
  - path: src/excel_vba_mcp/domain/statuses.py
    targets: [{type: module, name: operation and verification statuses}]
  - path: src/excel_vba_mcp/application/session_service.py
    targets:
      - {type: class, name: SessionService}
      - {type: method, name: recover_session}
      - {type: method, name: close_session}
  - path: src/excel_vba_mcp/application/compile_service.py
    targets: [{type: class, name: CompileService}]
  - path: src/excel_vba_mcp/application/macro_service.py
    targets: [{type: class, name: MacroService}]
  - path: src/excel_vba_mcp/security/audit.py
    targets: [{type: class, name: AuditLogger}]
  - path: src/excel_vba_mcp/storage/artifact_store.py
    targets:
      - {type: class, name: ArtifactStore}
      - {type: method, name: save_artifact}
      - {type: method, name: list_artifacts}
changes:
  - action: harden_error_and_recovery_paths
    description: Согласовать статусы, ошибки COM/UIA и cleanup/recovery для существующих services.
  - action: add_artifact_storage
    description: Добавить разрешённое хранилище артефактов с manifest, path checks и audit metadata.
  - action: add_fault_injection_tests
    description: Проверить worker failure, modal dialog timeout, Excel crash/closed workbook and cleanup branches.
acceptance_criteria:
  - каждый operation result имеет стабильный status и machine-readable error code при ошибке
  - cleanup освобождает только принадлежащий серверу Excel instance
  - artifact paths не выходят из разрешённого root
  - логи не содержат credentials, secrets или полного содержимого книги без необходимости
  - COM/UIA timeout и потеря Excel не оставляют сессию ложно успешной
  - audit trail фиксирует mutating operations и approvals
tests_to_run: ["unit fault-injection tests", "integration tests for Excel crash/timeout and cleanup", "artifact path policy tests"]
agents:
  - researcher: required; compare compatible recovery and artifact persistence practices
  - auditor: required; inspect COM lifecycle, error propagation and security boundaries
  - planner: required; design fault-injection matrix and evidence retention checks
  - validator: required_after_research_and_planner; independently review cleanup and artifact safety
verification_mode: dedicated_gate
```

## task_035 — Implement declarative test runner

```yaml
id: task_035
title: Implement bounded declarative workbook test runner
priority: critical
depends_on: [task_009, task_014, task_015, task_020, task_022, task_025, task_027, task_029, task_031, task_032, task_034]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.5. Testing & Snapshots; 8.5. TestRunner; 8.6. TestSpec; 12.2. Acceptance criteria A09"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A09 — Test Runner", "Bounded test specification and assertions"]
description: >
  Реализовать ограниченный декларативный формат тестовых сценариев, который координирует
  snapshot, разрешённые изменения объектов, compile, macro execution, assertions и artifacts.
  Не исполнять произвольный Python, shell, динамический COM или VBA-код из определения теста.
goals:
  - определить Pydantic модели TestSuite/TestCase/Step/Assertion
  - поддержать только операции и assertions, перечисленные в Спецификации
  - создавать отдельную disposable workbook copy для каждого тест-кейса
  - выполнять compile и macro steps через существующие application services
  - возвращать step-level и case-level structured statuses
  - сохранять run manifest, logs и artifacts через ArtifactStore
  - предоставлять MCP tools запуска и получения результатов тестирования
  - регистрировать read-only resource результата/сводки тестового запуска
extra_context: |
  Формат теста ограничен allowlisted actions. Поддерживаемые assertions включают macro status,
  sheet_exists, table row/column checks и targeted read-only output checks, заданные спецификацией.
  Тест-кейс не получает прямого доступа к Python interpreter, произвольным путям, COM objects,
  shell или динамическому выполнению кода. Каждый кейс стартует с собственного snapshot/copy;
  ошибка одного кейса не должна загрязнять следующий. Compile failure блокирует macro execution
  этого кейса. Внешние side effects макросов не откатываются snapshot-механизмом.
files:
  - path: src/excel_vba_mcp/domain/test_spec.py
    targets:
      - {type: model_group, name: TestSuite/TestCase/TestStep/Assertion DTOs}
      - {type: validator, name: allowed action and assertion validation}
  - path: src/excel_vba_mcp/application/test_runner.py
    targets:
      - {type: class, name: TestRunner}
      - {type: method, name: run_suite}
      - {type: method, name: run_case}
      - {type: method, name: execute_step}
      - {type: method, name: evaluate_assertion}
  - path: src/excel_vba_mcp/api/test_tools.py
    targets:
      - {type: function, name: test_run_suite}
      - {type: function, name: test_get_run}
  - path: src/excel_vba_mcp/storage/run_repository.py
    targets:
      - {type: class, name: RunRepository}
      - {type: method, name: save_run}
      - {type: method, name: get_run}
  - path: src/excel_vba_mcp/server.py
    targets: [{type: resource_registration, name: test run summary resource}]
changes:
  - action: add_bounded_test_schema
    description: Определить валидируемую декларативную схему без произвольного исполняемого кода.
  - action: orchestrate_existing_services
    description: Оркестрировать snapshot, compile, macro, Sheets, Tables, Names и assertions через существующие services.
  - action: persist_test_run
    description: Сохранять статус, шаги, диагностику и ссылки на artifacts в RunRepository/ArtifactStore.
acceptance_criteria:
  - неподдерживаемые actions/assertions отвергаются до запуска Excel operation
  - каждый test case выполняется в своей disposable copy
  - compile failure не приводит к запуску макроса
  - ошибка одного test case не влияет на baseline следующего case
  - statuses distinguish passed, failed, blocked, skipped and indeterminate as defined by domain statuses
  - результаты и artifacts доступны агенту через MCP tools/resource
  - A09 is satisfied without arbitrary Python/shell/COM execution
tests_to_run: ["unit tests for test schema/step state machine", "integration tests for multi-case suite isolation", "test runner MCP contract tests"]
agents:
  - researcher: required; review declarative test runner patterns and bounded assertion model
  - auditor: required; inspect execution boundary and test-case isolation
  - planner: required; design orchestration state machine and fixture suite
  - validator: required_after_research_and_planner; independently review injection resistance and acceptance coverage
verification_mode: dedicated_gate
```

## task_036 — Verify end-to-end test execution and isolation

```yaml
id: task_036
title: Verify end-to-end test runner workflows and case isolation
priority: critical
depends_on: [task_035]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "7.5. Testing & Snapshots; 12.1. Test fixtures; 12.2. Acceptance criteria A09"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A09 — Test Runner", "CleanProject.xlsm and TableNamesProject.xlsm fixtures"]
description: >
  Провести независимую интеграционную проверку полного сценария: disposable copy, разрешённое
  изменение тестовых данных/настроек, compile, macro execution, проверки выходов и сохранение
  evidence. Проверить изоляцию нескольких кейсов и отказные сценарии.
goals:
  - выполнить успешный сценарий на CleanProject.xlsm
  - выполнить сценарий с табличными настройками на TableNamesProject.xlsm
  - проверить compile failure и blocked macro step
  - проверить runtime error и сохранение diagnostic evidence
  - доказать, что каждый кейс получает отдельную копию и следующий кейс не наследует мутации
  - проверить resource retrieval тестового результата
extra_context: |
  Тестировать фактический Excel/VBA runtime в интерактивной сессии. Не считать snapshot гарантией
  отката внешних side effects. Любая тестовая макрокоманда, способная затронуть внешнюю систему,
  должна быть явно исключена fixture или разрешена в тестовой политике; не моделировать внешние
  транзакции, которых нет в Спецификации.
files:
  - path: tests/integration_excel/test_end_to_end_test_runner.py
    targets: [{type: test_module, name: End-to-end test runner verification}]
  - path: tests/fixtures/CleanProject.xlsm
    targets: [{type: fixture, name: CleanProject}]
  - path: tests/fixtures/TableNamesProject.xlsm
    targets: [{type: fixture, name: TableNamesProject}]
changes:
  - action: add_end_to_end_regression_suite
    description: Проверить реальные workflows и integration boundaries тестового runner.
acceptance_criteria:
  - successful suite returns passed with persisted artifacts
  - compile error blocks macro execution and returns compile diagnostics
  - runtime failure is represented as failed/indeterminate according to collected evidence
  - cases are isolated; failed case does not contaminate the next baseline
  - original fixture files remain unchanged
  - A09 passes on the target Excel build
tests_to_run: ["test_end_to_end_test_runner.py on target Windows/Excel"]
agents:
  - auditor: required; inspect end-to-end failure and cross-case contamination paths
  - validator: required; independently review workflow evidence and isolation claims
verification_mode: dedicated_verification
```

## task_037 — Verify Stage 1 acceptance and release readiness

```yaml
id: task_037
title: Verify Stage 1 acceptance matrix and release readiness
priority: critical
depends_on: [task_036]
source_reference: excel_vba_mcp_technical_spec.docx
source_section: "12.2. Acceptance criteria A01–A15; 13.2. Этапы реализации; 14. Definition of Done"
source_file: excel_vba_mcp_technical_spec.docx
source_blocks: ["A01–A15 — Acceptance criteria", "Stage 1 Definition of Done"]
description: >
  Выполнить заключительную независимую проверку Stage 1 по acceptance matrix, зафиксировать
  поддерживаемую конфигурацию Excel/Python, результаты регрессии, известные ограничения и
  воспроизводимые инструкции локального запуска. Новые функции в рамках release gate не добавлять.
goals:
  - проверить каждое применимое A01–A15 по наблюдаемому evidence
  - выполнить полный unit и integration regression на целевой конфигурации
  - проверить чистую установку из lockfile и STDIO startup
  - подтвердить отсутствие незапланированных Stage 2 функций
  - проверить README по настройке AccessVBOM, macro policy, запуску и ограничениям
  - сформировать release evidence и известные ограничения
extra_context: |
  Нельзя объявлять compile success без положительного сигнала от VBE compile flow. Нельзя
  приравнивать обнаружение ActiveX event procedure к доказательству работающего control click.
  Нельзя заявлять поддержку xls/xlsb/xlam или иных Office builds сверх фактически проверенных
  fixture/target configurations. Если xlam указан в Спецификации как возможный формат, в release
  report его поддержка должна быть обозначена только при наличии отдельной проверки.
files:
  - path: README.md
    targets:
      - {type: section, name: supported environment and installation}
      - {type: section, name: Trust Center and macro execution setup}
      - {type: section, name: limitations and recovery}
  - path: tests/
    targets: [{type: test_suite, name: complete Stage 1 regression suite}]
  - path: docs/release-evidence.md
    targets: [{type: document, name: Stage 1 acceptance and compatibility evidence}]
changes:
  - action: add_release_verification_evidence
    description: Зафиксировать результаты A01–A15, версии среды, fixture outcomes и известные ограничения.
acceptance_criteria:
  - every applicable A01–A15 has pass/fail/not-applicable evidence with reason
  - clean setup and MCP STDIO startup are reproducible
  - all required unit and integration tests pass on the declared target environment
  - original fixture files remain unchanged
  - no Stage 2 feature is included in the Stage 1 release
  - limitations and unsupported configurations are explicit
tests_to_run: ["complete unit and integration regression suite on declared target environment", "clean install/startup smoke test"]
agents:
  - auditor: required; audit acceptance evidence and scope compliance
  - planner: required; validate release checklist completeness
  - validator: required; independent final release gate
verification_mode: dedicated_release_gate
```

## 6. Execution and handoff rules

1. Выполнять задачи в порядке DAG, а не в порядке расположения разделов Спецификации. Зависимая задача не начинается до выполнения acceptance criteria своих прямых зависимостей.
2. Для `task_002`–`task_006` использовать результаты PoC как Go/No-Go. Если целевая Office-конфигурация не подтверждает требуемую операцию, остановить зависимую ветку и зафиксировать результат; не подменять её неподтверждённым механизмом.
3. Для каждой задачи с `dedicated_gate` / `dedicated_verification` сначала выполнить указанный agent review, затем реализацию/проверку. Агент-исполнитель не считается независимым Validator.
4. Для задач с `verification_mode: inline` исполнитель запускает `tests_to_run` и проверяет все `acceptance_criteria` до завершения задачи. Отдельную verification task не создавать.
5. Разрешать параллельную работу только над независимыми файлами/ветками. COM/Excel integration execution всегда сериализовать через одну STA worker и одну целевую Excel session.
6. Не объединять конфликтующие альтернативы. Любое решение, требующее выхода за Спецификацию, возвращать Product Owner как change request, а не добавлять в реализации автоматически.
7. После каждого mutation-capable task проверить baseline/snapshot, approval, audit trail и сохранность оригиналов. Не запускать unknown macro только ради теста.
8. Все semantic targets должны быть найдены по именам модулей, классов, методов, DTO и resource identifiers. Не использовать нестабильные позиционные/строчные координаты в качестве якорей.

## 7. Research references used for planning

- Official MCP Python SDK documentation: https://py.sdk.modelcontextprotocol.io/ — актуальный Python SDK interface and runtime requirements.
- Microsoft COM initialization: https://learn.microsoft.com/en-us/windows/win32/learnwin32/initializing-the-com-library — apartment-threaded COM initialization.
- Microsoft COM threading and apartments: https://learn.microsoft.com/en-us/windows/win32/com/processes--threads--and-apartments — STA ownership and apartment boundaries.
- Microsoft VBA Debug menu: https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/debug-menu — VBE Compile command behavior.
- pywinauto long-operation waits: https://pywinauto.readthedocs.io/en/uia/wait_long_operations.html — UIA wait/timeout practices for VBE dialogs.
- Excel Application.Run: https://learn.microsoft.com/en-us/office/vba/api/Excel.Application.Run — macro invocation and argument behavior.
- Excel Workbook.SaveCopyAs: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas — workbook copy semantics.
- Excel Workbook.VBProject: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.vbproject — VBA project object model entry point.
- Microsoft server-side Office automation considerations: https://support.microsoft.com/en-us/visio/considerations-for-server-side-automation-of-office — interactive-session and unattended automation limitations.

---

**Plan completion definition:** 37 semantic implementation/verification tasks are defined; their dependency DAG, risk gates, parallel work boundaries, semantic targets and validation conditions are explicit. This plan contains no code changes and does not authorize implementation choices beyond the validated specification.
