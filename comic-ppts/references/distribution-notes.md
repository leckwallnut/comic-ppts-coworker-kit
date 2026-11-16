# Distribution notes

This is the compressed distribution copy of the skill. The full-size originals were about 113 MB.

- **Opaque pictures** (backgrounds, hero concepts, cast sheets, the flattened ready-to-use backgrounds and the overview sheets, 53 files) are JPEG, quality 90, with the long edge capped at 1920 px. Their filenames now end in `.jpg`. The ready-to-use backgrounds had only about 3,000 faintly transparent edge pixels each, so they are stored as opaque JPEG.
- **Transparent panes, tabs, components and their SVGs** are unchanged, so the pixel coordinates in `layout-data.json` and `component-data.json` still apply.
- **`reference.pptx`** has the same slides, text, layouts and click animations. Only its embedded pictures were compressed, and identical pictures are stored once.
- **`assets/library-index.json`** lists the new `path`, `bytes` and `sha256` of every file, plus `originalSha256` for converted pictures.
