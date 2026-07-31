# Quick Links Pro

A SharePoint Framework (SPFx) web part that replaces the out-of-the-box SharePoint Quick Links with a more powerful, content-rich alternative, giving editors the flexibility to provide real context alongside every link.

![Quick Links Pro web part preview](https://www.camerongriffiths.com/images/quickLinksCustom.png)
![Quick Links Pro property pane](https://www.camerongriffiths.com/images/quickLinksCustom-pageProperties.png)

Part of the [SPFXWebparts](https://github.com/CamGriff/SPFXWebparts) collection. Full writeup and preview: [camerongriffiths.com/spfx/quicklinkspro](https://www.camerongriffiths.com/spfx/quicklinkspro)

## Features

- **Rich Tile Layout** — each link is displayed as a clean, rectangular tile with a title, extended description, and optional icon, arranged in a responsive two-column grid.
- **Extended Descriptions** — unlike the standard Quick Links web part, supports up to 400 characters of descriptive text per tile, giving content creators room to provide real context before a user clicks.
- **Fluent UI Icons** — choose from the full Microsoft Fluent UI icon library for a recognisable visual cue on each tile.
- **Fully Editor-Managed** — tiles are added, edited and removed directly from the property pane, no lists, no external data sources, no developer involvement.
- **Whole Tile Clickable** — the entire tile surface is a single clickable link, maximising the tap and click target on mobile and desktop alike.
- **Theme Aware** — tile borders and icons automatically adopt your SharePoint site's primary theme colour.
- **Bilingual Ready** — full English and French localisation included, with the architecture in place to support additional languages.

## Prerequisites

- SharePoint Framework development environment (Node.js, Heft toolchain)
- SharePoint Online tenant
- Site owner or SharePoint administrator permissions to deploy the web part

## Getting Started

```bash
npm install
npm run build
```

This produces the `.sppkg` file in `sharepoint/solution/`.

## Deployment

1. **Upload to the App Catalog** — upload the `.sppkg` to your tenant App Catalog. Choose tenant-wide or site-by-site deployment.
2. **Add to site** — go to Site Contents on the target site, Add an app, find the Quick Links Pro app, install.
3. **Add the web part** — add Quick Links Pro to a page via the web part picker.
4. **Configure** — open the property pane to add tiles, each with a title, description, icon, and link URL.

## Localisation

Ships with English and French out of the box. Additional languages can be added by extending the localisation resource files.

## License

MIT

## Author

[Cameron Griffiths](https://www.camerongriffiths.com), Microsoft 365 consultant based in Valencia, Spain.
