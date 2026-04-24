# XLStatUDF Website

Static SvelteKit foundation for the public XLStatUDF website.

## What it does

- renders Czech and English pages
- reads source documentation from the repository root `docs/`
- copies available installer `.exe` files into the website's static downloads folder

## Commands

```bash
npm install
npm run dev
npm run build
npm run deploy:aws
```

The documentation sync runs automatically before `dev` and `build`.

For AWS deployment details, see [`DEPLOY_AWS.md`](./DEPLOY_AWS.md).

Deployment versioning is driven by the repository root [`version.json`](../version.json).
