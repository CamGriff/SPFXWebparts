# ScrollToTop

A lightweight SharePoint Framework (SPFx) Application Customizer that adds a scroll-to-top button to SharePoint Online pages. A small addition that makes a big difference on long, content-rich intranet pages.

![ScrollToTop extension preview](https://www.camerongriffiths.com/images/ScrollToTop.png)

Part of the [SPFXWebparts](https://github.com/CamGriff/SPFXWebparts) collection. Full writeup and preview: [camerongriffiths.com/spfx/scrolltotop](https://www.camerongriffiths.com/spfx/scrolltotop)

## Features

- **Always There When You Need It** — sits quietly in the bottom right corner of every page, only appearing after the user has scrolled down, keeping the interface clean on shorter pages.
- **Smooth Scroll** — one click and the page glides back to the top instantly, no jarring jumps.
- **Theme Aware** — automatically adopts your SharePoint site's primary theme colour, no styling configuration required.
- **Site-Level Control** — installed once by the site owner, activates automatically on every page of the site. No page-by-page setup, no editor involvement needed.
- **Zero Footprint** — built as a lightweight SPFx Application Customizer with no external dependencies, fast to load, invisible until needed.
- **Bilingual Ready** — fully compatible with multilingual SharePoint sites.

## Prerequisites

- SharePoint Framework development environment (Node.js, Heft toolchain)
- SharePoint Online tenant
- Site owner or SharePoint administrator permissions to deploy the extension

## Getting Started

```bash
npm install
npm run build
```

This produces the `.sppkg` file in `sharepoint/solution/`.

## Deployment

1. **Upload to the App Catalog** — upload the `.sppkg` to your tenant App Catalog. Choose tenant-wide or site-by-site deployment.
2. **Add to site** — go to Site Contents on the target site, Add an app, find the ScrollToTop app, install.

That's it. As an Application Customizer, it activates automatically across every page of the site once installed, no further page-level configuration is required.

## Localisation

Fully compatible with multilingual SharePoint sites.

## License

MIT

## Author

[Cameron Griffiths](https://www.camerongriffiths.com), Microsoft 365 consultant based in Valencia, Spain.
