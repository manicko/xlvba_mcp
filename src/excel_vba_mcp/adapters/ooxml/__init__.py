"""Non-COM, streaming, bounded, read-only OOXML adapters (WP-E1).

These adapters read worksheet rows, shared strings and workbook metadata from an
immutable snapshot ``.xlsx`` / ``.xlsm`` file **without** occupying the single-STA
COM worker and **without** materializing the whole sheet in memory. They never
touch Excel/COM and never import ``pywin32``: the only input is a path to a file
already captured by the snapshot store (a ``Workbook.SaveCopyAs`` image).

The package is intentionally import-light so importing it never requires Windows.
Submodules:

* :mod:`~excel_vba_mcp.adapters.ooxml.zip_safety` — ZIP-bomb / traversal guards.
* :mod:`~excel_vba_mcp.adapters.ooxml.shared_strings` — streaming sharedStrings.
* :mod:`~excel_vba_mcp.adapters.ooxml.worksheet` — streaming worksheet rows.
* :mod:`~excel_vba_mcp.adapters.ooxml.workbook_meta` — workbook/rels metadata.
* :mod:`~excel_vba_mcp.adapters.ooxml.pivotcache` — saved PivotCache records.
"""