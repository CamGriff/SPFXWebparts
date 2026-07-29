# PageHeader

A custom SharePoint Framework (SPFx) web part for creating branded, configurable page headers in SharePoint Online. Replaces the default SharePoint header with a polished, fully configurable, on-brand experience.

![PageHeader web part preview](https://www.camerongriffiths.com/images/pageHeader.png)

Part of the [SPFXWebparts](https://github.com/CamGriff/SPFXWebparts) collection. Full writeup and preview: [camerongriffiths.com/spfx/pageheader](https://www.camerongriffiths.com/spfx/pageheader)

## Features

- **Custom Background Image** — upload any image directly from SharePoint via the built-in file picker, turning the header into a visual statement of your organisation's identity.
- **Smart Search Box** — a prominently placed, clearly visible search box with customisable placeholder text, connected to your SharePoint search results page.
- **Dynamic Breadcrumbs** — automatically displays the folder path from your Site Pages library on non-home pages, giving users instant context of where they are in the site structure.
- **Site Identity Tag** — a configurable site name label displayed on the header, linking back to the site home page.
- **Spotlight Button** — a flexible call-to-action button for pinning timely or high-priority content (campaigns, announcements, events), toggleable directly from the property pane.
- **Bilingual Ready** — full English and French localisation out of the box, with the architecture in place to support additional languages.
- **Theme Aware** — automatically inherits your SharePoint site theme colours, no additional configuration required.

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

1. **Upload to the App Catalog** — upload the `.sppkg` to your tenant App Catalog (`/sites/appcatalog/AppCatalog`). Choose tenant-wide or site-by-site deployment depending on how many sites you want to use the header on.
2. **Add to site** — go to Site Contents on the target site, Add an app, find the PageHeader app, install.
3. **Add the web part** — add PageHeader to a page via the web part picker, typically placed at the very top of the page.
4. **Configure** — open the property pane to set your background image, search placeholder text, site identity tag, and spotlight button content.

## Localisation

Ships with English and French out of the box. Additional languages can be added by extending the localisation resource files.

## License

MIT

## Author

[Cameron Griffiths](https://www.camerongriffiths.com), Microsoft 365 consultant based in Valencia, Spain.
