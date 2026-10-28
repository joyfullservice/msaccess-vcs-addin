# office_links_zoom

Regression fixture for Id-addable-only controls that cannot be created blank:

- **Zoom** combo (Type 4, Id 1733) — Access-managed list; export must not emit
  an empty `List` block.
- **Office Links** split MRU popup (Type 14, Id 2598) — children serialized when
  the submenu is customized; import adds them after `Add(14, 2598)`.
