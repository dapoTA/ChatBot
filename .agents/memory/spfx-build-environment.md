---
name: SPFx build environment
description: Runtime and configuration constraints for reliably building the pinned SharePoint Framework solution.
---

Build the SPFx 1.18 solution with Node 18, even when the main application uses a
newer Node release. Keep the SPFx build configuration isolated from root-level
ESM tooling configuration.

**Why:** SPFx 1.18 rejects Node 20, and its legacy PostCSS loader can walk into
the main application's ESM PostCSS configuration and fail with an ESM/CommonJS
loader error.

**How to apply:** Run SPFx install, bundle, and package steps under Node 18 and
keep an SPFx-local CommonJS PostCSS configuration. Do not downgrade the main
application runtime solely to build SharePoint assets.