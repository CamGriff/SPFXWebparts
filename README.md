# SPFXWebparts

A collection of SharePoint Framework (SPFx) web parts and extensions for SharePoint Online, built for real intranet governance, branding, and communication needs across client and personal projects.

Each solution lives in its own folder with a dedicated README covering features, prerequisites, and deployment steps. Full write-ups, screenshots, and demos for every solution are also published on [camerongriffiths.com/spfx](https://www.camerongriffiths.com/en/spfx).

## Solutions

| Solution | Type | Description |
|---|---|---|
| [PageHeader](./header-webpart) | Web Part | Branded, fully configurable page header with custom background, smart search box, dynamic breadcrumbs, and a spotlight call-to-action button. |
| [PageFooter](./pagefooter-webpart) | Web Part | List-driven, editor-managed page footer with grouped link columns, company logo, and social media icons. |
| [ScrollToTop](./scrolltotop-extension) | Extension | Lightweight Application Customizer that adds a theme-aware scroll-to-top button to every page on a site. |
| [QuickLinksPro](./quicklinkspro-webpart) | Web Part | Rich tile-based quick links, with extended descriptions and Fluent UI icons, editor-managed entirely from the property pane. |
| [CustomisableToolsMenu](./customisableToolsMenu) | Web Part (x2) | Two connected web parts giving each user a personalised, drag-and-drop launcher of their most-used tools on the intranet homepage. |
| [MessageBanner](./messageBanner) | Web Part | List-driven, severity-coded announcement banner with automatic expiration and an animated carousel for multiple active messages. |
| [TranslationDriftDashboard](./translationdriftdashboard-webpart) | Web Part | Single-site proof of concept surfacing translation drift (In Sync / Stale / Missing) across a multilingual SharePoint site, paired with companion PnP PowerShell scripts. |
| [PnPMultilingualGovernance](./pnp-multilingual-governance) | SPFx Solution | Tenant-wide successor to the Translation Drift Dashboard: multi-site drift detection, a translator nudge workflow, and Microsoft Graph email notifications, all in a single `.sppkg` with no Azure dependency. |

> Folder names above match the repo as of this writing. `PageFooter`, `ScrollToTop`, `QuickLinksPro`, and `TranslationDriftDashboard` are still being pushed, folder names are best-guess placeholders until confirmed, adjust the links once each is committed if the naming differs.

## Common Prerequisites

- SharePoint Framework development environment (Node.js, Heft toolchain)
- SharePoint Online tenant
- Site owner or SharePoint administrator permissions to deploy (tenant administrator for PnPMultilingualGovernance, which requires Graph API consent)

Each solution's own README has the exact build and deployment steps, since a couple require additional setup (a companion SharePoint list, PnP.PowerShell scripts, or Graph permissions).

## Getting Started

Clone the repo, then work within whichever solution folder you need:

```bash
git clone https://github.com/CamGriff/SPFXWebparts.git
cd SPFXWebparts/<solution-folder>
npm install
npm run build
```

Each folder produces its own `.sppkg` in `sharepoint/solution/`, deployed independently to your tenant App Catalog.

## License

MIT, see [LICENSE](./LICENSE). Free to use, adapt, and deploy in your own SharePoint environment.

## Author

[Cameron Griffiths](https://www.camerongriffiths.com) is a Microsoft 365 consultant based in Valencia, Spain, specialising in SharePoint Online, SPFx, PnP PowerShell, Power Platform, and Copilot Studio.

Bugs, questions, or ideas for additional solutions are welcome via [issues](https://github.com/CamGriff/SPFXWebparts/issues).
