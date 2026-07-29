# Customisable Tools Menu

A SharePoint Framework (SPFx) solution consisting of two connected web parts, giving each user a personalised launcher of their most-used tools directly on the intranet homepage.

The solution is driven by two SharePoint lists: a centrally managed tools catalogue maintained by administrators, and a per-user preferences list recording which tools each person has selected and in what order.

![Customisable Tools Menu, Tool Picker web part](https://www.camerongriffiths.com/images/toolsPicker.png)
![Customisable Tools Menu, Tool Display web part](https://www.camerongriffiths.com/images/toolsDisplay.png)

Part of the [SPFXWebparts](https://github.com/CamGriff/SPFXWebparts) collection. Full writeup and preview: [camerongriffiths.com/spfx/customisabletoolsmenu](https://www.camerongriffiths.com/spfx/customisabletoolsmenu)

## Components

- **Tool Picker Web Part** — lives on a dedicated settings page, allowing users to build their own selection of tools from the catalogue. Tools can be added or removed with a single click, and reordered using drag and drop. Selections save automatically back to the preferences list.
- **Tool Display Web Part** — sits on the intranet homepage and renders each user's personal toolkit as a clean, icon-driven grid, fully personalised so every user sees their own selection, in their chosen order, every time they visit.

## Features

- **Configurable Maximum** — supports a configurable maximum of eight tools per user, keeping the homepage focused and uncluttered.
- **Adjustable Grid Columns** — the number of columns in the grid is adjustable via the property pane, no code changes required.
- **Bilingual Ready** — fully multilingual, automatically detecting the user's SharePoint language preference and serving content in English or French accordingly.
- **Brand Customisable** — styling is customisable to align with your organisation's branding.

## Prerequisites

- SharePoint Framework development environment (Node.js, Heft toolchain)
- SharePoint Online tenant
- Site owner or SharePoint administrator permissions to deploy the solution

## Getting Started

```bash
npm install
npm run build
```

This produces the `.sppkg` file in `sharepoint/solution/`.

## Deployment

1. **Upload to the App Catalog** — upload the `.sppkg` to your tenant App Catalog. Choose tenant-wide or site-by-site deployment.
2. **Provision the two lists** — the master tools catalogue and the per-user preferences list (see the source for column schema).
3. **Add to site** — go to Site Contents on the target site, Add an app, find the Customisable Tools Menu app, install.
4. **Add the web parts** — add the Tool Picker web part to a dedicated settings page, and the Tool Display web part to the homepage.
5. **Populate the catalogue** — administrators add the full list of available tools to the master tools list, each user then builds their own personal selection via the Tool Picker.

## Localisation

Automatically detects the user's SharePoint language preference and serves English or French accordingly.

## License

MIT

## Author

[Cameron Griffiths](https://www.camerongriffiths.com), Microsoft 365 consultant based in Valencia, Spain.
