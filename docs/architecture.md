# MTR Control Architecture

The project follows a layered architecture with explicit boundaries between business logic, orchestration, infrastructure, and delivery mechanisms.

## Layers

| Layer | Path | Responsibility |
| --- | --- | --- |
| Domain | `src/mtr_control/domain` | Enterprise rules, entities, value objects, events, and domain services. |
| Application | `src/mtr_control/application` | Use-case orchestration, commands, queries, and DTOs. |
| Infrastructure | `src/mtr_control/infrastructure` | Technical adapters such as persistence, configuration, logging, and third-party integrations. |
| Interfaces | `src/mtr_control/interfaces` | Entry points such as HTTP APIs, CLIs, and serialized schemas. |
| Shared | `src/mtr_control/shared` | Small cross-cutting primitives that do not belong to a single layer. |

## Boundary rules

- Domain code is framework-free and contains no infrastructure imports.
- Application code coordinates domain behavior through use cases.
- Infrastructure implements adapters required by application and domain ports.
- Interfaces translate external requests into application commands or queries.
