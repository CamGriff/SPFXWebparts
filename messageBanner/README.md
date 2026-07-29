# Message Banner

A SharePoint Framework (SPFx) web part providing a list-driven message banner for SharePoint Online, keeping employees informed with severity-coded, auto-expiring alerts managed entirely by a communications team from a SharePoint list.

Part of the [SPFXWebparts](https://github.com/CamGriff/SPFXWebparts) collection. Full writeup, preview and demo video: [camerongriffiths.com/spfx/messagebanner](https://www.camerongriffiths.com/spfx/messagebanner)

## Features

- **List-Driven Messages** — banners are stored and managed directly in a SharePoint list. The communications team adds, edits or deactivates items and the banner updates automatically, no deployments or developer involvement required.
- **Message Severity Levels** — each banner is assigned Low, Medium or High severity, colour-coded blue, orange and red respectively, so employees can immediately gauge urgency at a glance.
- **Automatic Expiration** — set an expiration date on any message and it automatically stops displaying once that date passes.
- **Animated Flip Carousel** — when multiple messages are active simultaneously, they cycle automatically with a smooth vertical flip transition, so all messages get visibility without cluttering the page.
- **See More Link** — optionally attach a hyperlink to any message with a fully customisable button label, or fall back to a sensible default.
- **Zero Footprint When Empty** — when there are no active messages, the web part collapses completely, leaving no empty space on the page.
- **Bilingual Ready** — full English and French localisation included, with language detected automatically based on the user's SharePoint interface language.

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
2. **Provision the messages list** — create the SharePoint list that drives the banner content, including severity, expiration date, and optional link fields (see the source for column schema).
3. **Add to site** — go to Site Contents on the target site, Add an app, find the Message Banner app, install.
4. **Add the web part** — add Message Banner to a page via the web part picker, typically near the top of the homepage.
5. **Manage messages** — the communications team adds, edits, and deactivates messages directly in the list, no further deployment needed for content changes.

## Localisation

Ships with English and French out of the box, detected automatically from the user's SharePoint interface language. Additional languages can be added by extending the localisation resource files.

## License

MIT

## Author

[Cameron Griffiths](https://www.camerongriffiths.com), Microsoft 365 consultant based in Valencia, Spain.
