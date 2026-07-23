# Architecture

MTR_Control uses a layered VBA architecture designed for maintainability and workbook portability.

## Layers

1. **UI Layer** (`src/UI`): Contains forms and presentation-only event handlers. UI modules must not contain business rules.
2. **Service Layer** (`src/Services`): Coordinates workflows, validation, repositories, and domain models.
3. **Repository Layer** (`src/Repository`): Encapsulates workbook access. Excel ranges must be read into arrays before processing.
4. **Domain Model Layer** (`src/Models`): Represents materials, movements, arrivals, and future business concepts.
5. **Core Layer** (`src/Core`): Provides shared application entry points, settings, logging, utility helpers, and error-handling conventions.

## Rules

- Every VBA module must include `Option Explicit`.
- Use repositories for workbook reads and writes.
- Use services to orchestrate business workflows.
- Keep domain objects free of workbook-specific behavior.
- Centralize diagnostics through `modLogger`.
- Handle errors consistently with module and procedure context.
- Avoid duplicated logic by extracting reusable helpers.
- Never iterate directly over worksheet cells for business processing; load ranges into arrays first.

## Current Scope

Release `v0.1.0` is a foundation release. It provides the skeleton only and intentionally does not implement matching logic.
