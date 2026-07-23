# MTR_Control

MTR_Control is an enterprise-grade VBA project skeleton for material tracking and reconciliation workflows.

## Release

Current foundation release: `v0.1.0`.

## Project Layout

```text
src/
  Core/          Shared entry points, settings, logging, and utilities
  Models/        Domain model classes
  Repository/    Workbook persistence and range access abstractions
  Services/      Application service layer
  Reports/       Report generation modules
  UI/            Forms and presentation-only code
docs/            Architecture and project documentation
tests/           Test assets and test modules
sample_data/     Sample workbooks or CSV data for local validation
releases/        Release notes and packaged exports
```

## Architecture Principles

- Use `Option Explicit` in every VBA module.
- Keep business logic out of UI forms.
- Access workbook data through repositories.
- Coordinate workflows through services.
- Represent business concepts with domain models.
- Use centralized logging and consistent error handling.
- Avoid copy/paste implementation patterns.
- Read Excel ranges into arrays before processing.

## Getting Started

1. Import the files under `src/` into a macro-enabled Excel workbook.
2. Run `Main` from `src/Core/modMain.bas`.
3. Extend services and repositories before adding UI behaviors.

## Status

This release intentionally provides only the foundation skeleton. Matching logic is not implemented yet.
