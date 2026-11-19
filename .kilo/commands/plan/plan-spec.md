# Objective

Transform a validated implementation specification into a dependency-aware implementation execution plan composed of independently executable semantic task specifications.

---
# Constraints

- ONLY produce an implementation plan.
- NEVER redesign architecture.
- NEVER invent requirements not supported by the specification.
- NEVER use line numbers.
- ALWAYS use semantic targets (function names, class names, variable names, module paths, etc.).
- Prefer incremental evolution.
- Preserve existing architecture.
- Prefer extending existing modules over introducing new abstractions.
- Avoid speculative implementation work.
- Optimize for implementation sequencing, reviewability, and parallel execution.

The implementation plan MUST NOT mirror the conceptual task list from the specification.  
Instead, reorganize work into implementation-ready execution tasks optimized for:

- dependencies
- rollout order
- architectural boundaries
- validation
- implementation safety

---

# Workflow

1. **Study the specification**  
   Study ONLY the specification and files explicitly provided by the user.  
   Do NOT inspect unrelated documents.

2. Decompose the plan into logical execution blocks
- Identify dependencies and execution order
- Assess implementation, rollout, regression, and compatibility risks
- Identify multiple implementation paths alternatives when relevant, for maintainability, future evolution, and project conventions
- **Not choose the final implementation approach when technical uncertainty exists**
- Keep the scope minimal and avoid speculative redesign
- For each block, determine whether the following agents are required:
  - **High risks** - all agents below 
  - **Auditor** — deeper code/architecture investigation due to uncertainty or complexity
  - **Researcher** — modern best practices, multiple viable approaches, architectural/support implications
  - **Planner** — detailed pre-implementation design, architecture, testing, or complex execution
  - **Validator** — independent plan review when implementation risk is high
**Important: Never change any code or existing audit phase files.**


3. Ask user as the Product Owner
Skip this step unless there are business uncertainty.
Do not ask about technical implementation or architecture - this is your task.

Questions should focus on:
- product behavior and usage examples
- business rules
- user expectations or experience
.
Provide questions with options to choose basing on the best practices and highlight recommended options. 

Collect answers for the the specification.


3. For every risky task:
   - First launch a **Researcher** agent to identify the best implementation approaches compatible with the current architecture, with an eye toward future maintainability and extensibility.
   - Then launch a **Validator** agent to critically review the proposed approach against the current architecture and long-term maintainability.

4. **Handle multiple implementation approaches**  
   If multiple viable approaches exist, launch a **Researcher** agent to identify modern best practices that fit the current architecture.  
   Research is mandatory whenever:
   - multiple architectural options exist
   - framework best practices influence the implementation
   - external libraries or patterns are involved
   - scalability or maintainability may be affected

5. **Build the execution DAG**  
   Identify:
   - implementation dependencies
   - rollout order
   - isolated implementation units
   - parallel execution groups

6. **Decompose execution blocks into implementation tasks**  
   Each task must represent one independently implementable unit.  
   Avoid:
   - mixed infrastructure + feature work
   - mixed implementation + testing
   - oversized implementation tasks

7. **Define semantic implementation targets**  
   For every task specify:
   - affected modules
   - affected classes
   - affected services
   - affected functions
   - semantic insertion points

   Never reference line numbers — they change. Semantic anchors (function names, class names, variable names, etc.) are stable.

8. **Prepare task specifications**  
   Base every task on the template:  
   # For simple tasks (single function, trivial/small effort, low risk):
# verification is INLINE — implementor runs tests_to_run and checks
# acceptance_criteria before marking the task complete.
# Do NOT create a separate verification task.

id: task_template

title: Add audit logging for user creation

priority: high

depends_on:
  - task_003
  - task_007

# ── Plan/Audit report source (REQUIRED when task originates from a plan or audit) ──
# source_reference: \.ai\plans\PLAN_XX.md
# source_section: TX — Task Title
# source_file: .kilo/commands/audit/...           # if extracting from existing file
# source_blocks:                                  # if extracting specific blocks
#   - "BLOCK N — Block Name"

description: >
  Add audit logging after successful user creation.

goals:
  - create audit log record
  - avoid affecting transaction flow
  - preserve backward compatibility

# ── Extra context from plan (verbatim, do not summarize) ──
# Include full checklists, YAML frontmatter, file paths, agent instructions,
# runtime steps, anti-patterns — everything from the plan's XML <details> block.
# extra_context:
#  

files:
  - path: app/services/user_service.py

    targets:
      - type: class
        name: UserService

      - type: method
        name: create_user

    semantic_anchors:

      insert_after:
        type: function_call
        value: validate_email

      insert_before:
        type: return_statement

changes:

  - action: add_code

    description: >
      Call audit_service.log_user_created()

    code_hint: |
      audit_service.log_user_created(
          user_id=user.id,
          created_by=admin.id
      )

acceptance_criteria:
  - audit log created
  - no transaction rollback issues
  - tests pass


 
   Group all tasks into a single Markdown plan.

9. **Insert test tasks**  
   Add test tasks only for non-trivial features.  
   Tests must:
   - validate user-visible behavior
   - exercise real workflows
   - detect regressions
   - cover integration boundaries

10. **Insert verification tasks**  
    Add dedicated verification tasks for multi-stage or high-risk changes.

---

# Conflict Resolution

Prefer in this order:

1. Safety
2. Existing architecture
3. Higher-confidence evidence
4. Incremental implementation

Never merge conflicting implementation approaches into a single task.  
Surface conflicts explicitly.

---

# Output

Generate the plan:

```
plan.md
```

The plan must represent the optimal implementation execution sequence, not the order presented in the original specification.
