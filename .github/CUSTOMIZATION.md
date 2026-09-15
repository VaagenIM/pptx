# Embed configuration

Set these variables on the embedding page to customize the viewer:

```css
:root {
  --pptx-primary: #2563eb;
  --pptx-primary-strong: #1d4ed8;
  --pptx-primary-ink: #ffffff;
  --pptx-primary-wash: rgb(108 158 248 / 13%);
  --pptx-text-color: #f4f2ec;
  --pptx-text-soft: #b4b6bd;
  --pptx-text-muted: #797e89;
  --pptx-logo-color: #2563eb;
  --pptx-font-family: Inter, system-ui, sans-serif;
  --pptx-mono-font-family: ui-monospace, monospace;
  --pptx-font-size-body: 1rem;
  --pptx-font-size-meta: 0.78rem;
  --pptx-font-size-control: 0.68rem;
  --pptx-font-size-heading: 1.35rem;
  --pptx-font-size-brand: 0.9rem;
  --pptx-surface-deep: #0b0d12;
  --pptx-background: #10131a;
  --pptx-surface: #151922;
  --pptx-surface-raised: #1b202a;
  --pptx-surface-hover: #292f3b;
  --pptx-border: #303744;
  --pptx-border-soft: #252b35;
  --pptx-topbar-background: rgb(17 20 27 / 94%);
  --pptx-stage-background: #10131a;
  --pptx-sidebar-background: #151922;
  --pptx-dialog-background: #151922;
  --pptx-dialog-shadow: 0 24px 70px rgb(0 0 0 / 38%);
  --pptx-shadow: 0 24px 70px rgb(0 0 0 / 38%);
  --pptx-danger: #ff9a92;
  --pptx-fullscreen-background: #08090c;
  --pptx-slide-background: #000;
}
```

These variables control the embedded viewer’s colors, typography, font sizes, surfaces,
borders, shadows, logo, and fullscreen backgrounds. Unset variables use the viewer defaults.

## Required embed script

Place [`pptx-embed.min.js`](test/pptx-embed.min.js) after the iframe elements:

```html
<script src="pptx-embed.min.js"></script>
```

The script finds every `.pptx-embed` iframe, reads the variables from the parent page, and
forwards them to the viewer. It also watches the parent document root for inline variable
changes and forwards updates without reloading the iframe. This is required for cross-origin
embeds because parent-page CSS cannot directly style the iframe document.
