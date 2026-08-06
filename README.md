# Site Content Index

An SPFx web part for KM/SPO oversight. Points at one or more SharePoint sites and automatically lists every list and library on each, no manual entry required per list, just add the site URL and it stays current as content changes.

## What it does

- Scans each configured site and displays its lists/libraries: name, type, item count, last modified date, and whether permissions are unique or inherited
- Groups results by site (collapsible) or shows a flat sortable table
- Live search filters by name across whatever's currently loaded
- Optional filter to show only libraries with unique (broken-inheritance) permissions, useful for spotting governance drift
- Optional "stale" flag on any library not modified within a configurable number of days
- Optional title, shown/hidden and custom text, both configurable
- Fully theme-reactive: colors follow whatever theme is applied to the site, no hardcoded palette

## Configuration

Open the web part's edit panel:

| Setting | What it does |
|---|---|
| Target Sites | Paste a site URL, press Enter to add it as a chip. Add as many sites as needed. |
| Include system lists | Off by default. Turn on to show SharePoint's default plumbing lists (Site Assets, Style Library, etc.) alongside real content. |
| Group results by site | On by default. Turn off for one flat table across all sites. |
| Show only libraries with unique permissions | Filters the whole view down to just the lists/libraries that have broken inheritance. |
| Flag libraries not modified in (days) | Enter a number to flag stale libraries. Leave blank or 0 to disable. |
| Show title | Toggle the heading on/off. |
| Title text | Optional custom heading. Leave blank for the default "Site Content Index." |

## Permissions

Only shows what the logged-in user already has access to see. There's no elevated or service-account access, if you lose access to a site, or someone with less access views the page, results reflect their own permissions, not yours.

## Deployment

Deployed via site collection app catalog. Repo: `spfx-metadata-compliance-dashboard`'s sibling project, cloned from the Recognition Feed scaffold. Solution name: `site-content-index`.

## Tech

SPFx 1.23.2, React, PnPjs (`@pnp/sp`). No third-party UI libraries. Custom property pane control (`SiteEntryChipInputField`) built using SPFx's documented `IPropertyPaneField` pattern, not the newer `PropertyPaneCustomField` helper, which isn't exported in this SPFx version.