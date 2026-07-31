# Recognition Feed

An SPFx web part that displays recent award winners as a scrolling card feed. Built for a SharePoint landing page as a lightweight way to keep recognition visible without a separate dashboard page nobody visits.

## What it does

- Reads winners from a SharePoint list (built and tested against a Critical Cog / Employee of the Month list, but not hardcoded to it)
- Shows the most recent N months of winners as auto-scrolling cards, pause/play control included
- Pulls each winner's photo from their Microsoft 365 profile, no photo library to maintain
- Click any card for the full detail: justification, team members if it's a team award
- A "Nominate Someone" tile lives in the same scroll, linking out to the nomination form

## Why it's dynamic instead of hardcoded

Every field is mapped through the property pane instead of assumed. Name, photo source, award type, month, year, justification, team fields, all of it is picked from real dropdowns populated from whatever list you point the web part at, using [@pnp/spfx-property-controls](https://github.com/pnp/sp-dev-fx-property-controls) for the list and column pickers.

That means the same web part works on a list with a completely different schema than the one it was built against. Point it at a different list, remap the fields, done. No code changes.

## Tech stack

- SPFx 1.23, React function components, TypeScript
- PnPjs for all SharePoint calls
- `@pnp/spfx-property-controls` for the list/column picker property pane
- No Graph calls, no external CDN dependencies, same-origin SharePoint REST only

## Built entirely through CI, no local Node install

This one's worth calling out. It was built on a locked-down government machine with no ability to install Node.js or Git locally. The whole project, scaffolding, every file, every build, went through a GitLab CI/CD pipeline running on Army's DevSecOps Platform (DSOP). Code gets written and committed through GitLab's Web IDE, a pipeline job installs dependencies and runs the SPFx build in a container, and the finished `.sppkg` comes back as a downloadable pipeline artifact.

No local dev environment exists for this project at all. That's not a limitation, it's how it was built from the first commit.

## Status

Compiles and packages clean through the pipeline. Not yet deployed against a real site with real data, the site collection app catalog for the target site isn't enabled yet. Once that's in place, this section gets updated with actual results instead of expected behavior.

## Local development

There is no local dev environment for this project. To build:

```
npm install
npm run build
```

Run through CI, not locally. Produces a `.sppkg` file in `sharepoint/solution`, downloadable as a pipeline artifact.