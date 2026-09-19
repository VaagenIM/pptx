# pptx

Served at <https://pptx.iktim.no>.

The site provides browsable PowerPoint files and embeddable presentations for use in other projects.

PowerPoint embed links use deterministic UUIDv5 IDs. The `data` branch publishes
`uuid_map.json`, which maps each ID to its source presentation path, while
compressed files are stored as `pptx/<uuid>.pptx`. The Pages site contains both
the generated directory index and the original presentations. The publishing
workflow carries the map forward and uses Git rename history, so moving or
renaming a presentation keeps its ID.

See [CUSTOMIZATION.md](CUSTOMIZATION.md) for embedding and customization instructions.

## Credits

The browser viewer is powered by
[pptx-web](https://github.com/costinEEST/pptx-web) and
[`pptx-renderer`](https://github.com/aiden0z/pptx-renderer).
`pptx-web` is available under the MIT license, and `pptx-renderer` under the
Apache-2.0 license.
