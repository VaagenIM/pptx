# pptx

Served at <https://pptx.iktim.no>.

The site provides browsable PowerPoint files and embeddable presentations for use in other projects.

PowerPoint embed links use deterministic UUIDv5 IDs. The public branch publishes
`map.json`, which maps each ID to its compressed presentation path; the directory
index reads the current file list from `main`.

See [CUSTOMIZATION.md](CUSTOMIZATION.md) for embedding and customization instructions.

## Credits

The browser viewer is powered by
[costinEEST/pptx-web](https://github.com/costinEEST/pptx-web), under its MIT license.
