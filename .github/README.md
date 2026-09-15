# pptx

Public-facing PowerPoint files.

## Browse the files

The repository has a GitHub Pages site at <https://vaagenim.github.io/pptx/>. It indexes the
entire `main` branch and provides search, folder navigation, PowerPoint filtering, downloads,
and buttons for copying file paths or raw URLs.

The site is deployed automatically by `workflows/pages.yml` after pushes to `main`.
In repository settings, set **Pages → Build and deployment → Source** to **GitHub Actions**
if Pages has not been enabled yet.

PowerPoint files can be opened in the self-hosted browser viewer from the **View online**
shortcut. The viewer is provided by the pinned `viewer` submodule and is built into the
Pages artifact at `/pptx/viewer/`; files are rendered in the browser and are not uploaded.
