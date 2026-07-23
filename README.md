# MTR Control

MTR Control is structured as a layered application to keep business rules independent from delivery and infrastructure details.

## Architecture

```text
src/mtr_control/
├── domain/          # Core business entities, domain services, and events
├── application/     # Use cases, commands, queries, and data transfer objects
├── infrastructure/  # Configuration, persistence adapters, integrations, and logging
├── interfaces/      # API, CLI, and external schemas
└── shared/          # Cross-cutting helpers shared across layers
tests/
├── unit/            # Fast tests for isolated units
└── integration/     # Tests for adapters and end-to-end flows
docs/                # Architecture and operational documentation
scripts/             # Developer and automation scripts
```

## Dependency rule

Dependencies should point inward:

```text
interfaces -> application -> domain
infrastructure -> application/domain
```

The `domain` package must not import from `application`, `infrastructure`, or `interfaces`.

## Development

Install the project in editable mode:

```bash
python -m pip install -e .
```

Run tests:

```bash
python -m pytest
```
