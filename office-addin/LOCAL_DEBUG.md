# Local Office Debug

This workflow runs the Office.js custom functions locally and registers local manifests for desktop Excel on Windows.

## What It Does

- creates and trusts a local `localhost` development certificate in the current user store
- builds the add-in into `office-addin/dist/`
- generates `office-addin/dist/manifest.local.xml` with `https://localhost:3000/...` URLs
- serves the built files over local HTTPS
- refreshes the trusted shared-folder catalog at `\\localhost\EvalyticsOfficeAddin`

## VS Code Buttons

Use these from `Run and Debug`:

- `Evalytics: Reset Debug Session`: closes Excel, rebuilds, refreshes the catalog, starts the local server, and launches Excel without clearing the Office WEF cache.
- `Evalytics: Hard Reset Debug Session`: does the same, but also clears the Office WEF cache. Use it after changing `functions.json`, manifest identity, or custom-functions metadata.

## Typical Flow

1. Run `Evalytics: Reset Debug Session`.
2. If the add-ins were already added before, Excel should keep them available.
3. After a hard reset, go to `Home > Add-ins > Advanced > Shared Folder` and add both `Evalytics` and `Evalytics Docs` again.
4. Use `=VERSION()` to confirm the Office.js custom-functions runtime is loaded.

## Notes

- The local manifest is Windows desktop sideloading oriented.
- The normal reset deliberately keeps the Office WEF cache so Excel does not forget manually added add-ins on every iteration.
- The manifest keeps production documentation/support URLs, but runtime asset URLs point to localhost.
- Desktop Excel discovers XML manifests through a trusted shared-folder catalog. The helper script creates a local catalog at `office-addin/.catalog` and registers `\\localhost\EvalyticsOfficeAddin` as trusted.
- Some Windows policies block `New-SmbShare` even from PowerShell. In that case, share `office-addin/.catalog` manually in File Explorer as `EvalyticsOfficeAddin`, then restart Excel.
- If a function keeps returning an old metadata-related error, run `Evalytics: Hard Reset Debug Session`.
