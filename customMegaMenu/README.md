# Custom Mega Menu

A SharePoint Framework (SPFx) Application Customizer that adds a list-driven mega menu to the top of every page on a SharePoint Online site, with up to three levels of navigation managed entirely by editors from a SharePoint list.

Part of the [SPFXWebparts](https://github.com/CamGriff/SPFXWebparts) collection.

## Features

- **List-Driven Navigation**: menu items live in a `MegaMenuItems` SharePoint list. Editors add, reorder, or remove links and the menu updates on the next page load, no redeployment needed.
- **Three-Level Hierarchy**: top-level items sit in the menu bar. Hovering one opens a flyout with its second-level links, each with its own third-level links indented beneath it.
- **Overlay Flyouts**: flyouts float above page content instead of pushing the page down, and stay open while the pointer moves into them.
- **Custom Ordering**: a `SortOrder` column controls the order of items at every level.
- **Open in New Tab**: any item can be set to open in a new tab (with `noopener noreferrer`).
- **Flyout-Only Parents**: a top-level item can be left without a URL and act purely as a heading for its flyout.
- **Fails Safe**: if the list is missing or can't be read, the menu renders empty and logs to the console instead of breaking the page.

## Prerequisites

- SharePoint Framework development environment (Node.js 22, Heft toolchain)
- SharePoint Online tenant
- Site owner or SharePoint administrator permissions to deploy the extension

## The `MegaMenuItems` List

Create a custom list named **`MegaMenuItems`** on each site where the menu runs. The menu reads from the current site, so each site needs its own copy. The columns below use internal names:

| Column | Type | Notes |
|---|---|---|
| `Title` | Single line of text | Link text shown in the menu. |
| `NavUrl` | Hyperlink | Link target. Leave empty for a flyout-only parent. |
| `ParentId` | Lookup (to `MegaMenuItems`, `ID`) | Parent item. Empty = top-level item. |
| `Levels` | Choice | Records the item's level for editors. The menu builds its tree from `ParentId`, not this column. |
| `SortOrder` | Number | Ascending order among siblings. |
| `OpenInNewTab` | Yes/No | Opens the link in a new tab. |

Up to 500 items are loaded.

## Getting Started

```bash
npm install
npm run build
```

This produces `custom-megamenu.sppkg` in `sharepoint/solution/`.

To debug against a live site, set your tenant domain and start the dev server:

```powershell
$env:SPFX_SERVE_TENANT_DOMAIN = "<tenant>.sharepoint.com"
npm run start
```

Update `pageUrl` in `config/serve.json` to point at a real page on your site.

## Deployment

1. **Upload to the App Catalog**: upload the `.sppkg` to your tenant App Catalog. Choose whether to make it available to all sites.
2. **Provision the list**: create `MegaMenuItems` on the target site with the columns above.
3. **Activate on a site**: go to Site Contents on the target site, add the app, and install it. For tenant-wide deployment, the extension is registered in the Tenant Wide Extensions list instead.
4. **Manage the menu**: editors maintain links directly in the list. Content changes need no further deployment.

## License

MIT

## Author

[Cameron Griffiths](https://www.camerongriffiths.com), Microsoft 365 consultant based in Valencia, Spain.
