# Evalytics Office.js Add-in

This folder contains the cross-platform Office.js custom-functions implementation of Evalytics.

## Scope

- Excel on Windows, Mac, and the web
- statistical worksheet functions implemented in TypeScript
- no legacy output-formatting layer

## Build

```powershell
npm install
npm run build
```

Build output is written to `dist/`.

## Local Excel Debug

For Windows desktop sideloading and local HTTPS hosting, see `LOCAL_DEBUG.md`.

## In-Excel Documentation

The manifest exposes an `Evalytics` ribbon group with a `Function Docs` button. It opens `taskpane.html`, a searchable side panel generated from the same `functions.json` metadata that drives Excel custom-function help.

## Current Status

Implemented in the new runtime:

- random generators and normal interval probability
- single-column random array generators: `GENERATE.NORM.ARRAY`, `GENERATE.INT.ARRAY`
- weighted descriptive statistics
- filtered percentiles
- Shapiro-Wilk and Kolmogorov-Smirnov normality checks
- one-sample, paired, and two-sample tests
- Spearman correlation
- chi-square goodness-of-fit
- one-way ANOVA
- basic two-argument `FILL` value repetition
- locale-aware numeric text parsing with `PARSE.NUMBER`

Still pending for full parity with the archived Excel-DNA version:

- repeated-measures ANOVA
- ANCOVA
- contingency analyses
- correlation matrix builder
- pivot family
