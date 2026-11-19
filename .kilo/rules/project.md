
1. **English Only**  
   All comments, logs, docstrings, error messages, and documentation must be in clear English.

2. **Production Code is King**  
   If tests conflict with architecture or business logic — fix or remove the tests. Never distort production code for tests.

3. **Strict Separation of Concerns**  
   Enforce clear boundaries between UI, API, business logic, and data layers.

4. **Single Responsibility**  
   Modules, services, components, and functions must be small and focused on one thing.

5. **Avoid Overengineering**  
   Prefer simple, obvious, and maintainable solutions over complex abstractions.

6. **Composition Over Inheritance**  
   Prefer composition in both backend and frontend.

7. **Follow Existing Patterns**  
    Respect and follow established patterns and conventions in the codebase. Do not introduce new abstractions without strong justification.   

8. **Meaningful & Consistent Naming**  
    Use clear, concise, and consistent names for files, functions, components, and variables.

9. **VBA Identifier Safety**  
   VBA identifiers are case-insensitive. Before adding or renaming any variable, constant, procedure, parameter, type, or other identifier in VBA, verify that the chosen name (in any letter case or abbreviation form) does not collide with a VBA reserved keyword, statement, built-in function, type, or object name (for example, `eNum` is interpreted as `Enum`). Use a distinct, non-reserved name to avoid compile-time syntax errors.

10. **VBA Declaration Placement**  
    In VBA modules, all `Option` statements, `Const`, `Dim`/`Public`/`Private` module-level declarations, `Type`, and `Enum` blocks must appear in the module's declarations section before the first `Sub`, `Function`, or `Property` procedure. Placing `Type`/`Enum` definitions after procedures causes compile errors (e.g. "User-defined type not defined") and must be avoided.

17. **Workspace Cleanliness**  
    Before creating any file, ask: 1) Is there an exact place for this in the task or module architecture? 2) Is it needed for project operation and future maintenance? If both are no — write only to `.ai/tmp`.





