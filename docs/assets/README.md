---
id: xlvba-assets
domain: assets
tags:
  - documentation
  - diagrams
  - assets
related:
  - xlvba-architecture
---

# Documentation assets

This directory holds diagram sources for the `xlvba` documentation.

Diagrams are maintained as **Mermaid source files** (`.mmd`) here and are
embedded **inline** in the docs as fenced `mermaid` code blocks. The `.mmd` file
is the canonical source; the inline block in the doc must match it. When a
diagram changes, update the `.mmd` file and the inline copy together.

Mermaid renders natively in most Markdown viewers (including the Kilo document
viewer and GitHub). No build step or image export is required.

## Files

| File | Diagram type | Embedded in |
|------|--------------|-------------|
| `architecture-overview.mmd` | `flowchart` | [Architecture](../04-architecture.md) — "High-level architecture diagram" |
| `request-lifecycle.mmd` | `sequenceDiagram` | [Architecture](../04-architecture.md) — "Request lifecycle" |

## Conventions

- Keep node and participant identifiers stable; they mirror the layer names in
  [Architecture](../04-architecture.md).
- Use relative links from this directory (`../04-architecture.md`).
- Do not add generated images (`.png`, `.svg`); the source of truth is the
  `.mmd` file.